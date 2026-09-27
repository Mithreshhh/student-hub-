const path = require('path');
const crypto = require('crypto');
const { BlobServiceClient } = require('@azure/storage-blob');
const { DefaultAzureCredential } = require('@azure/identity');

// Containers must be created in the storage account beforehand; this service never creates them.
const CONTAINERS = Object.freeze({
  RESUMES: 'resumes',
  CERTIFICATES: 'certificates',
  ACTIVITY_IMAGES: 'activity-images'
});
const KNOWN_CONTAINERS = new Set(Object.values(CONTAINERS));

// Extensions we keep on blob names, and the Content-Type each blob is served with.
// The Content-Type comes from this allowlist, never from the client-supplied MIME string.
const CONTENT_TYPES = Object.freeze({
  '.pdf': 'application/pdf',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.png': 'image/png'
});

class BlobStorageError extends Error {
  constructor(message, cause) {
    super(message);
    this.name = 'BlobStorageError';
    this.code = 'BLOB_STORAGE_ERROR';
    if (cause) this.cause = cause;
  }
}

let blobServiceClient = null;

const getAccountName = () => {
  const accountName = (process.env.AZURE_STORAGE_ACCOUNT_NAME || '').trim();
  if (!accountName) {
    throw new BlobStorageError(
      'AZURE_STORAGE_ACCOUNT_NAME is not set. Add it to backend/.env (see .env.example) to enable file uploads.'
    );
  }
  return accountName;
};

const getAccountUrl = () => `https://${getAccountName()}.blob.core.windows.net`;

// Lazily created so the server can boot (and serve non-upload routes) without Azure configured.
// DefaultAzureCredential uses Azure CLI / VS Code / env credentials locally and Managed Identity in Azure.
const getBlobServiceClient = () => {
  if (!blobServiceClient) {
    blobServiceClient = new BlobServiceClient(getAccountUrl(), new DefaultAzureCredential());
  }
  return blobServiceClient;
};

// Translate SDK failures into messages that say what to fix.
const toBlobStorageError = (error, containerName) => {
  if (error instanceof BlobStorageError) return error;

  if (error.name === 'CredentialUnavailableError' || error.name === 'AuthenticationError' || error.name === 'AggregateAuthenticationError') {
    return new BlobStorageError(
      'Azure authentication failed. Locally, sign in with `az login`; in Azure App Service, enable Managed Identity. ' +
      `Details: ${error.message}`,
      error
    );
  }
  if (error.statusCode === 403) {
    return new BlobStorageError(
      `Access to storage account "${getAccountName()}" was denied (${error.code || 403}). ` +
      'The signed-in identity needs the "Storage Blob Data Contributor" role on the account or container.',
      error
    );
  }
  if (error.statusCode === 404 && error.code === 'ContainerNotFound') {
    return new BlobStorageError(`Blob container "${containerName}" does not exist in storage account "${getAccountName()}".`, error);
  }
  if (error.code === 'ENOTFOUND') {
    return new BlobStorageError(`Storage account "${getAccountName()}" could not be reached. Check AZURE_STORAGE_ACCOUNT_NAME.`, error);
  }
  return new BlobStorageError(`Blob storage request failed: ${error.message}`, error);
};

const assertKnownContainer = (containerName) => {
  if (!KNOWN_CONTAINERS.has(containerName)) {
    throw new BlobStorageError(`Unknown blob container "${containerName}"`);
  }
};

// Build "<uuid><ext>" and pick a Content-Type. The original filename is only used for its
// extension, and only when that extension is on the allowlist; otherwise fall back to the MIME type.
const describeFile = (file) => {
  let extension = path.extname(file.originalname || '').toLowerCase();
  if (!CONTENT_TYPES[extension]) {
    extension = Object.keys(CONTENT_TYPES).find((ext) => CONTENT_TYPES[ext] === file.mimetype) || '';
  }
  return {
    blobName: `${crypto.randomUUID()}${extension}`,
    contentType: CONTENT_TYPES[extension] || 'application/octet-stream'
  };
};

/**
 * Upload a Multer memoryStorage file to a container.
 * @returns {Promise<string>} the blob's URL, suitable for storing in PostgreSQL
 */
const uploadBlob = async (containerName, file) => {
  assertKnownContainer(containerName);
  if (!file || !Buffer.isBuffer(file.buffer)) {
    throw new BlobStorageError('uploadBlob expects a Multer memoryStorage file with a buffer');
  }

  const { blobName, contentType } = describeFile(file);
  try {
    const blockBlobClient = getBlobServiceClient()
      .getContainerClient(containerName)
      .getBlockBlobClient(blobName);

    await blockBlobClient.uploadData(file.buffer, {
      blobHTTPHeaders: { blobContentType: contentType }
    });
    return blockBlobClient.url;
  } catch (error) {
    throw toBlobStorageError(error, containerName);
  }
};

/**
 * Parse a stored URL into { containerName, blobName }.
 * Returns null for anything that is not a blob in this app's storage account and containers
 * (for example legacy "/uploads/..." paths), so callers can skip it safely.
 */
const parseBlobUrl = (blobUrl) => {
  if (!blobUrl || typeof blobUrl !== 'string') return null;

  let url;
  try {
    url = new URL(blobUrl);
  } catch {
    return null;
  }

  let accountHost;
  try {
    accountHost = new URL(getAccountUrl()).host;
  } catch {
    return null;
  }
  if (url.host.toLowerCase() !== accountHost.toLowerCase()) return null;

  const [containerName, ...rest] = url.pathname.replace(/^\//, '').split('/');
  const blobName = decodeURIComponent(rest.join('/'));
  if (!KNOWN_CONTAINERS.has(containerName) || !blobName) return null;

  return { containerName, blobName };
};

/**
 * Delete the blob behind a stored URL. Never throws: failures are logged and reported as false,
 * so a cleanup problem cannot undo or block the database change that triggered it.
 * @returns {Promise<boolean>} true if a blob was deleted
 */
const deleteBlob = async (blobUrl) => {
  if (!blobUrl) return false;

  const parsed = parseBlobUrl(blobUrl);
  if (!parsed) {
    console.warn(`Skipping blob deletion for URL outside the configured storage account: ${blobUrl}`);
    return false;
  }

  try {
    const response = await getBlobServiceClient()
      .getContainerClient(parsed.containerName)
      .getBlobClient(parsed.blobName)
      .deleteIfExists({ deleteSnapshots: 'include' });
    return response.succeeded;
  } catch (error) {
    console.error('Blob deletion failed:', toBlobStorageError(error, parsed.containerName).message);
    return false;
  }
};

// Delete several blobs, ignoring empty values. Never throws.
const deleteBlobs = (blobUrls) => Promise.all(blobUrls.filter(Boolean).map(deleteBlob));

/**
 * Startup check: confirm the account name is set and a token can be obtained.
 * Only logs; it does not stop the server, since only upload routes depend on storage.
 */
const verifyBlobStorageConfig = async () => {
  try {
    const accountName = getAccountName();
    await new DefaultAzureCredential().getToken('https://storage.azure.com/.default');
    console.log(`🗄️  Azure Blob Storage ready (account: ${accountName})`);
  } catch (error) {
    console.warn(`⚠️  Azure Blob Storage is not usable, file uploads will fail: ${toBlobStorageError(error).message}`);
  }
};

module.exports = {
  CONTAINERS,
  BlobStorageError,
  uploadBlob,
  deleteBlob,
  deleteBlobs,
  parseBlobUrl,
  verifyBlobStorageConfig
};
