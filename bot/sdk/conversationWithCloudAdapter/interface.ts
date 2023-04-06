// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CloudAdapter, ConversationReference } from "botbuilder";
import {
  NotificationTargetStorage,
  BotSsoConfig,
  CommandOptions,
  CardActionOptions,
} from "../conversation/interface";

/**
 * Options to initialize {@link NotificationBot}.
 */
export interface NotificationOptions {
  /**
   * An optional input of bot app Id.
   *
   * @remarks
   * If `botAppId` is not provided, `process.env.BOT_ID` will be used by default.
   */
  botAppId?: string;
  /**
   * An optional store to persist bot notification connections.
   * 
   * @remarks
   * If `store` is not provided, a default conversation reference store will be used,
   * which stores notification connections into:
   *  - `.notification.localstore.json` if running locally
   *  - `${process.env.TEMP}/.notification.localstore.json` if `process.env.RUNNING_ON_AZURE` is set to "1"
   * 
   * It's recommended to use your own store for production environment.
   */
  store?: ConversationReferenceStore;
  /**
   * An optional storage to persist bot notification connections.
   *
   * @deprecated Use `store` to customize the way to persist bot notification connections instead.
   * @remarks
   * If `storage` is not provided, a default local file storage will be used,
   * which stores notification connections into:
   *   - `.notification.localstore.json` if running locally
   *   - `${process.env.TEMP}/.notification.localstore.json` if `process.env.RUNNING_ON_AZURE` is set to "1"
   *
   * It's recommended to use your own shared storage for production environment.
   */
  storage?: NotificationTargetStorage;
}

/**
 * Options to initialize {@link ConversationBot}
 */
export interface ConversationOptions {
  /**
   * The bot adapter. If not provided, a default adapter will be created:
   * - with `adapterConfig` as constructor parameter.
   * - with a default error handler that logs error to console, sends trace activity, and sends error message to user.
   *
   * @remarks
   * If neither `adapter` nor `adapterConfig` is provided, will use BOT_ID and BOT_PASSWORD from environment variables.
   */
  adapter?: CloudAdapter;

  /**
   * If `adapter` is not provided, this `adapterConfig` will be passed to the new `CloudAdapter` when created internally.
   *
   * @remarks
   * If neither `adapter` nor `adapterConfig` is provided, will use BOT_ID and BOT_PASSWORD from environment variables.
   */
  adapterConfig?: { [key: string]: unknown };

  /**
   * Configurations for sso command bot
   */
  ssoConfig?: BotSsoConfig;

  /**
   * The command part.
   */
  command?: CommandOptions & {
    /**
     * Whether to enable command or not.
     */
    enabled?: boolean;
  };

  /**
   * The notification part.
   */
  notification?: NotificationOptions & {
    /**
     * Whether to enable notification or not.
     */
    enabled?: boolean;
  };

  /**
   * The adaptive card action handler part.
   */
  cardAction?: CardActionOptions & {
    /**
     * Whether to enable adaptive card actions or not.
     */
    enabled?: boolean;
  };
}

export interface ConversationReferenceStore {
  /**
   * Add a conversation reference to the store. If overwrite, update existing one, otherwise add when not exist.
   * @returns true if added or updated, false if not changed.
   */
  add(
    key: string,
    reference: Partial<ConversationReference>,
    options?: ConversationReferenceStoreAddOptions
  ): Promise<boolean>;

  /**
   * Remove a conversation reference from the store.
   * @returns true if exist and removed, false if not changed.
   */
  remove(
    key: string,
    reference: Partial<ConversationReference>
  ): Promise<boolean>;

  /**
   * List stored conversation reference by page.
   * @returns a paged list.
   */
  list(
    pageSize?: number,
    continuationToken?: string
  ): Promise<PagedData<Partial<ConversationReference>[]>>;
}

export interface ConversationReferenceStoreAddOptions {
  overwrite?: boolean;
}

export interface PagedData<T> {
  data: T;
  continuationToken?: string;
}
