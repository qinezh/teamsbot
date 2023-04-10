import { ConversationReference } from "botbuilder";
import { Container, CosmosClient } from "@azure/cosmos";

import {
    ConversationReferenceStore,
    ConversationReferenceStoreAddOptions,
    PagedData
} from "../sdk/conversationWithCloudAdapter/interface";

export class CosmosStore implements ConversationReferenceStore {
    private initializePromise?: Promise<Container>;
    private readonly client: CosmosClient;
    private readonly databaseName: string;
    private readonly containerName: string;

    constructor(
        connectionString: string,
        databaseName: string,
        containerName: string,
    ) {
        this.client = new CosmosClient(connectionString);
        this.databaseName = databaseName;
        this.containerName = containerName;
    }

    public async add(
        key: string,
        reference: Partial<ConversationReference>,
        options?: ConversationReferenceStoreAddOptions
    ): Promise<boolean> {
        const container = await this.initialize();
        if (options?.overwrite === false && await this.exists(key)) {
            return false;
        }

        const item = Object.assign({}, reference, { id: key });
        await container.items.upsert(item);
        return true;
    }

    public async remove(key: string, reference: Partial<ConversationReference>): Promise<boolean> {
        const container = await this.initialize();

        if (await this.exists(key) === false) {
            return false;
        }

        await container.item(key, key).delete();
        return true;
    }

    public async list(pageSize?: number, continuationToken?: string): Promise<PagedData<Partial<ConversationReference>>> {
        const container = await this.initialize();
        const querySpec = {
            query: 'SELECT * FROM c',
            parameters: [],
        };

        const items = [];
        do {
            const { resources, continuation } = await container.items
                .query(querySpec, { continuationToken, maxItemCount: pageSize })
                .fetchAll();

            items.push(...resources);
            continuationToken = continuation;
        } while (continuationToken);

        return {
            data: items,
            continuationToken,
        }
    }

    // Get user's conversation reference by user's AAD object id
    public async getUserByUserId(userId: string): Promise<Partial<ConversationReference>> {
        const container = await this.initialize();
        const querySpec = {
            query: 'SELECT * FROM c WHERE c.user.aadObjectId = @userId AND c.conversation.conversationType = "personal"',
            parameters: [{ name: '@userId', value: userId }],
        };
        const { resources } = await container.items.query(querySpec).fetchAll();
        return resources.length > 0 ? resources[0] : undefined;
    }

    // Get user's conversation reference by user's email
    public async getUserByUserEmail(userEmail: string): Promise<Partial<ConversationReference>> {
        const userId = await this.getUserIdByUserEmail(userEmail);
        return userId ? await this.getUserByUserId(userId) : undefined;
    }

    // Get user's AAD object id by user's email from another service, e.g. Graph API
    private async getUserIdByUserEmail(userEmail: string): Promise<string> {
        throw new Error("Method not implemented.");
    }        

    // Initialize to create container if not exists yet
    private async initialize(): Promise<Container> {
        if (!this.initializePromise) {
            const { database } = (await this.client.databases.createIfNotExists({ id: this.databaseName }));
            const { container } = (await database.containers.createIfNotExists({ id: this.containerName }));

            this.initializePromise = Promise.resolve(container);
        }

        return this.initializePromise;
    }

    private async exists(key: string): Promise<boolean> {
        const container = await this.initialize();
        const { resource: item } = await container.item(key, key).read();
        return !!item;
    }
}