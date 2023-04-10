import { BotBuilderCloudAdapter } from "../../sdk";
import ConversationBot = BotBuilderCloudAdapter.ConversationBot;
import config from "./config";
import { BlobStore } from "../blobStore";
import { CosmosStore } from "../cosmosStore";

// export const blobStore = new BlobStore(config.blobConnectionString, config.blobContainerName);
export const cosmosStore = new CosmosStore(
  config.cosmosConnectionString,
  config.cosmosDatabaseName,
  config.cosmosContainerName,
);

// Create bot.
export const bot = new ConversationBot({
  // The bot id and password to create CloudAdapter.
  // See https://aka.ms/about-bot-adapter to learn more about adapters.
  adapterConfig: {
    MicrosoftAppId: config.botId,
    MicrosoftAppPassword: config.botPassword,
    MicrosoftAppType: "MultiTenant",
  },
  // Enable notification
  notification: {
    enabled: true,
    store: cosmosStore,
  },
});
