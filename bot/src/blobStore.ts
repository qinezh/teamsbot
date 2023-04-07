import { ConversationReference } from "botbuilder";
import { ConversationReferenceStore, ConversationReferenceStoreAddOptions, PagedData } from "../sdk/conversationWithCloudAdapter/interface";
import { ContainerClient, ContainerListBlobFlatSegmentResponse } from "@azure/storage-blob";

// A sample implementation to use Azure Blob Storage as notification target storage
export class BlobStore implements ConversationReferenceStore {
    private readonly client: ContainerClient;
    private initializePromise?: Promise<unknown>;

    // This implementation uses connection string and container name to connect Azure Blob Storage
    constructor(connectionString: string, containerName: string) {
        this.client = new ContainerClient(connectionString, containerName);
    }

    async get(key: string): Promise<Partial<ConversationReference>> {
        await this.initialize();

        const blobName = this.normalizeKey(key);

        try {
            const stream = await this.client.getBlobClient(blobName).download();
            const content = await this.streamToBuffer(stream.readableStreamBody);
            return JSON.parse(content.toString());
        } catch (error) {
            if (error.statusCode === 404) {
                return undefined;
            } else {
                throw error;
            }
        }
    }

    async add(
        key: string,
        reference: Partial<ConversationReference>,
        options?: ConversationReferenceStoreAddOptions
    ): Promise<boolean> {
        await this.initialize();

        const blobName = this.normalizeKey(key);
        const content = JSON.stringify(reference);
        if (options.overwrite) {
            try {
                await this.client.getBlockBlobClient(blobName).upload(content, Buffer.byteLength(content));
                return true;
            } catch (error) {
                if (error.statusCode !== 404) {
                    throw error;
                }
            }
        } else {
            try {
                if (await this.get(key) === undefined) {
                    await this.client.getBlockBlobClient(blobName).upload(content, Buffer.byteLength(content));
                    return true;
                }
            } catch (error) {
                if (error.statusCode !== 404) {
                    throw error;
                }
            }
        }

        return false;
    }

    async remove(key: string, reference: Partial<ConversationReference>): Promise<boolean> {
        await this.initialize();

        const blobName = this.normalizeKey(key);

        try {
            await this.client.getBlobClient(blobName).delete();
            return true;
        } catch (error) {
            if (error.statusCode !== 404) {
                throw error;
            }

            return false;
        }
    }

    async list(pageSize?: number, continuationToken?: string): Promise<PagedData<Partial<ConversationReference>>> {
        await this.initialize();

        const result = new Array<Partial<ConversationReference>>();
        const iterator = this.client.listBlobsFlat().byPage({ maxPageSize: pageSize, continuationToken });
        const response: ContainerListBlobFlatSegmentResponse = (await iterator.next()).value;

        for (const blob of response.segment.blobItems) {
            try {
                const stream = await this.client.getBlockBlobClient(blob.name).download();
                const content = await this.streamToBuffer(stream.readableStreamBody);
                result.push(JSON.parse(content.toString()));
            } catch (error) {
                if (error.statusCode !== 404) {
                    throw error;
                }
            }
        }

        return {
            data: result,
            continuationToken: response.continuationToken
        };
    }

    // Initialize to create container if not exists yet
    private initialize(): Promise<unknown> {
        if (!this.initializePromise) {
            this.initializePromise = this.client.createIfNotExists();
        }

        return this.initializePromise;
    }

    // A help method to normalize key to meet Azure Blob naming requirement
    private normalizeKey(key: string): string {
        return encodeURIComponent(key);
    }

    // A helper method used to read a Node.js readable stream into a Buffer
    private streamToBuffer(stream: NodeJS.ReadableStream): Promise<Buffer> {
        return new Promise((resolve, reject) => {
            const chunks: Buffer[] = [];
            stream.on("data", (data) => {
                chunks.push(Buffer.isBuffer(data) ? data : Buffer.from(data));
            });
            stream.on("end", () => {
                resolve(Buffer.concat(chunks));
            });
            stream.on("error", reject);
        });
    }
}