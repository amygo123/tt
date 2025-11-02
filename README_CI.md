# GitHub Actions: Build & Release

这两个 workflow 让你“推到 GitHub 就自动构建”，打 tag 则自动产出发布包。

## 放置位置
把这两个文件放到仓库：
- `.github/workflows/build.yml`
- `.github/workflows/release.yml`

## 触发方式
- push 到 `main` 或 `master`：自动 **Build**，并把产物作为 **Artifact**（`StyleWatcher-win-x64`）上传。
- 打 tag（例如 `v1.0.0`）：自动 **Publish** 并创建 GitHub Release，附带 `StyleWatcher-win-x64.zip` 产物。

## 产物说明
- 构建目标：`.NET 8`、`win-x64`、**自包含 + 单文件**（release.yml 里显式指定）。
- 构建目录：`out/win-x64/`
- 发布包：`StyleWatcher-win-x64.zip`

## 常见问题
1) **我需要代码签名吗？**  
   目前 workflow 未包含代码签名步骤；如需签名可在 `release.yml` 中添加 `signtool` 或使用第三方签名 Action（需要上传证书到 GitHub Secrets）。

2) **NuGet 私有源/代理？**  
   在 `setup-dotnet` 之后增加 `nuget.config` 或 `dotnet nuget add source` 即可；必要时配置 Secrets（如 `NUGET_AUTH_TOKEN`）。

3) **分支不是 main/master？**  
   修改 `build.yml` 的 `on.push.branches` 和 `pull_request.branches` 即可。

## 打标签发布
```bash
git tag v1.0.0
git push origin v1.0.0
```

发布完成后，前往 GitHub Releases 页面下载 `StyleWatcher-win-x64.zip`。
