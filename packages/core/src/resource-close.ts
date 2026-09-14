export interface CloseableResource {
  close(): Promise<void>;
}

export interface ResourceCloseFailure<Resource extends CloseableResource> {
  readonly resource: Resource;
  readonly error: unknown;
}

export class OwnedResources<Resource extends CloseableResource> {
  readonly #resources = new Set<Resource>();
  #closing: Promise<readonly ResourceCloseFailure<Resource>[]> | undefined;

  constructor(resources: Iterable<Resource> = []) {
    for (const resource of resources) this.#resources.add(resource);
  }

  get size(): number {
    return this.#resources.size;
  }

  add(resource: Resource): void {
    this.#resources.add(resource);
  }

  close(): Promise<readonly ResourceCloseFailure<Resource>[]> {
    if (this.#closing !== undefined) return this.#closing;
    const resources = [...this.#resources];
    const closing = Promise.all(
      resources.map((resource) =>
        Promise.resolve()
          .then(() => resource.close())
          .then(
            (): readonly ResourceCloseFailure<Resource>[] => {
              this.#resources.delete(resource);
              return [];
            },
            (error: unknown): readonly ResourceCloseFailure<Resource>[] => [
              { resource, error },
            ],
          ),
      ),
    )
      .then((failures) => failures.flat())
      .finally(() => {
        this.#closing = undefined;
      });
    this.#closing = closing;
    return closing;
  }
}
