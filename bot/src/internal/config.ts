const config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  blobConnectionString: process.env.BLOB_CONNECTION_STRING,
  blobContainerName: process.env.BLOB_CONTAINER_NAME,
  cosmosConnectionString: process.env.COSMOS_CONNECTION_STRING,
  cosmosDatabaseName: process.env.COSMOS_DATABASE_NAME,
  cosmosContainerName: process.env.COSMOS_CONTAINER_NAME,
};

export default config;
