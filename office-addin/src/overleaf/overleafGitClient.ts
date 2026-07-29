import * as git from "isomorphic-git";
import type {
  FsClient,
  HttpClient,
  PromiseFsClient,
} from "isomorphic-git";

type WritableFsClient = PromiseFsClient & {
  promises: {
    mkdir(path: string): Promise<void>;
    stat(path: string): Promise<unknown>;
    writeFile(path: string, data: Uint8Array): Promise<void>;
  };
};

export type GitSyncProgress = "cloning" | "pulling" | "writing" | "pushing";

export interface PushPdfOptions {
  remoteUrl: string;
  filePath: string;
  token: string;
  data: Uint8Array;
  onProgress?: (progress: GitSyncProgress) => void;
}

export interface PushPdfResult {
  changed: boolean;
  branch: string;
  commitOid?: string;
}

export class OverleafGitClient {
  private readonly dir: string;
  private readonly fs: WritableFsClient;
  private readonly http: HttpClient;

  constructor(options: {
    dir: string;
    fs: FsClient;
    http: HttpClient;
  }) {
    this.dir = options.dir;
    this.fs = options.fs as unknown as WritableFsClient;
    this.http = options.http;
  }

  async pushPdf(options: PushPdfOptions): Promise<PushPdfResult> {
    const onAuth = () => ({ username: "git", password: options.token });

    if (await this.hasClone()) {
      options.onProgress?.("pulling");
      await git.pull({
        fs: this.fs,
        http: this.http,
        dir: this.dir,
        author: {
          name: "Slide2Pdf",
          email: "slide2pdf@localhost",
        },
        fastForwardOnly: true,
        singleBranch: true,
        onAuth,
      });
    } else {
      options.onProgress?.("cloning");
      await git.clone({
        fs: this.fs,
        http: this.http,
        dir: this.dir,
        url: options.remoteUrl,
        depth: 1,
        noTags: true,
        singleBranch: true,
        onAuth,
      });
    }

    const branch = await git.currentBranch({ fs: this.fs, dir: this.dir });
    if (!branch) {
      throw new Error("无法确定 Overleaf 项目的默认分支。");
    }
    const baseOid = await git.resolveRef({
      fs: this.fs,
      dir: this.dir,
      ref: "HEAD",
    });

    options.onProgress?.("writing");
    await this.createParentDirectories(options.filePath);
    await this.fs.promises.writeFile(
      `${this.dir}/${options.filePath}`,
      options.data,
    );

    await git.add({
      fs: this.fs,
      dir: this.dir,
      filepath: options.filePath,
    });
    const status = await git.status({
      fs: this.fs,
      dir: this.dir,
      filepath: options.filePath,
    });
    if (status === "unmodified") {
      return { changed: false, branch };
    }

    const commitOid = await git.commit({
      fs: this.fs,
      dir: this.dir,
      message: `Update ${options.filePath} from Slide2Pdf`,
      author: {
        name: "Slide2Pdf",
        email: "slide2pdf@localhost",
      },
    });

    options.onProgress?.("pushing");
    try {
      const pushResult = await git.push({
        fs: this.fs,
        http: this.http,
        dir: this.dir,
        ref: branch,
        force: false,
        onAuth,
      });
      const refErrors = Object.values(pushResult.refs)
        .filter((status) => !status.ok)
        .map((status) => status.error);
      if (!pushResult.ok || pushResult.error || refErrors.length) {
        throw new Error(
          [pushResult.error, ...refErrors].filter(Boolean).join("\n") ||
            "Overleaf 拒绝了 Git push。",
        );
      }
    } catch (error) {
      await this.restoreBase(branch, baseOid);
      throw error;
    }

    return { changed: true, branch, commitOid };
  }

  private async hasClone(): Promise<boolean> {
    try {
      await this.fs.promises.stat(`${this.dir}/.git`);
      return true;
    } catch {
      return false;
    }
  }

  private async restoreBase(branch: string, baseOid: string): Promise<void> {
    await git.writeRef({
      fs: this.fs,
      dir: this.dir,
      ref: `refs/heads/${branch}`,
      value: baseOid,
      force: true,
    });
    await git.checkout({
      fs: this.fs,
      dir: this.dir,
      ref: branch,
      force: true,
    });
  }

  private async createParentDirectories(filePath: string): Promise<void> {
    const directories = filePath.split("/").slice(0, -1);
    let currentPath = this.dir;

    for (const directory of directories) {
      if (!directory || directory === ".") continue;
      currentPath += `/${directory}`;
      try {
        await this.fs.promises.stat(currentPath);
      } catch {
        await this.fs.promises.mkdir(currentPath);
      }
    }
  }
}
