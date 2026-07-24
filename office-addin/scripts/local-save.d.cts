export interface LocalSaveRequest {
  slideKey: string;
  suggestedName: string;
  forceNewPath: boolean;
  data: Uint8Array;
}

export interface LocalPdfSaverDependencies {
  choosePath?: (suggestedName: string) => Promise<string>;
  writeFile?: (path: string, data: Uint8Array) => Promise<void>;
}

export class LocalPdfSaver {
  constructor(dependencies?: LocalPdfSaverDependencies);
  save(request: LocalSaveRequest): Promise<{ fileName: string }>;
}

export function createLocalSaveMiddleware(): (
  request: unknown,
  response: unknown,
  next: () => void,
) => Promise<void>;
