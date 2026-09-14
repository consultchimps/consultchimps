export class RetryableClose {
  readonly #release: () => Promise<void>;
  #open = true;
  #closing: Promise<void> | undefined;

  constructor(release: () => Promise<void>) {
    this.#release = release;
  }

  get isOpen(): boolean {
    return this.#open;
  }

  close(): Promise<void> {
    if (this.#closing !== undefined) return this.#closing;
    if (!this.#open) return Promise.resolve();
    const closing = Promise.resolve()
      .then(() => this.#release())
      .then(() => {
        this.#open = false;
      })
      .finally(() => {
        this.#closing = undefined;
      });
    this.#closing = closing;
    return closing;
  }
}
