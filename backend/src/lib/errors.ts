/** The item exists but the operation doesn't apply to it (a folder, or no preview/download available). → 400 */
export class UnsupportedItemError extends Error {}

/** The X-Confirm-Delete name doesn't match the item (stale list or forged request). → 409 */
export class DeleteNotConfirmedError extends Error {}

/** The uploaded file is not allowed (blocked type, or content doesn't match its extension). → 415 */
export class RejectedFileError extends Error {}
