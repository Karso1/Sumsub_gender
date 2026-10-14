# Sumsub Gender from Excel / 从 Excel 填写 Sumsub 性别

A macOS automation tool that reads user supplied gender values from an Excel workbook and fills **blank** document gender fields in Sumsub. It opens a visible Chrome session for manual sign in, then processes applicants by UID. Existing values are preserved. Preview mode does not save changes.

这是一个 macOS 自动化工具：从 Excel 读取用户提供的性别，根据 UID 查找 Sumsub 申请人，只填写**空白**的证件性别字段。程序会打开可见的 Chrome，由操作者自行登录。已有性别不会被覆盖；预览模式不会保存。

## Setup / 安装

Requires macOS, Node.js 18+, Google Chrome, and access to your own Sumsub client. Run `npm install` in this repository. Put one `.xlsx` file in `input/`; if there are several, the launcher asks you to choose. The workbook must have a `用户ID` or `UID` column and a `gender`, `性别`, or `真实性别` column. Supported values include `male`/`female`/`X`, `男`/`女`, and common abbreviations.

需要 macOS、Node.js 18+、Google Chrome，以及你自己的 Sumsub 客户权限。在仓库目录运行 `npm install`，把 `.xlsx` 放进 `input/`。如果有多个文件，启动时可输入序号选择。表格需包含 `用户ID` 或 `UID` 列，以及 `gender`、`性别` 或 `真实性别` 列。支持 `male`、`female`、`X`、`男`、`女` 及常见简写。

Set your Sumsub client ID before running:

```sh
export SUMSUB_CLIENT_ID='your-client-id'
npm run preview   # inspect first 10 rows; no changes
npm run apply     # save blank values and verify after reload
```

运行前请设置你自己的 Sumsub 客户 ID。也可以双击 `01-预览前10条.command` 或 `02-正式更新全部.command`；双击启动时需在终端环境中预先设置 `SUMSUB_CLIENT_ID`。登录 Chrome 后回到终端按回车。运行时程序会复用一个页面；请不要关闭该 Chrome 窗口。

Results and resume checkpoints are written under `output/<workbook-name>-<content-hash>/`. A repeat run of the same workbook skips previously saved or matching records. If a Sumsub value already exists and differs from Excel, the tool records a conflict and leaves it unchanged.

结果和续跑记录保存在 `output/<表格名>-<内容哈希>/`。同一表格再次运行会跳过已经保存或原本一致的记录。如果 Sumsub 已有性别与 Excel 不同，程序记录冲突并保留原值。

## Knowledge points / 主要知识点

- **Browser automation / 浏览器自动化**: Playwright controls Chrome with role based locators, waits for visible controls, clicks radio buttons, saves, and reloads to verify the result. It uses direct applicant URLs because Sumsub's single page routing can open a new tab instead of navigating the search page.
- **Spreadsheet parsing / 表格解析**: Node.js reads the `.xlsx` ZIP structure and XML worksheets, maps columns by header, and normalizes gender values without an Excel application dependency.
- **Idempotency and resume / 幂等与续跑**: Existing fields are left alone, successful rows are checkpointed, and each workbook content hash gets independent progress files.
- **Auditability / 审计记录**: Each processed UID is written to a CSV report with its outcome; errors and conflicting values are visible for manual review.
- **Data minimization / 数据最小化**: The workbook and generated reports stay local. This public repository excludes `input/`, `output/`, browser runtime files, and credentials through `.gitignore`.

## Privacy / 隐私

Do not commit real applicant spreadsheets, audit reports, or credentials. Sumsub gender is personal information; only run this tool for records you are authorized to edit. The visible Chrome session requires manual login and is closed at the end of a run.

不要提交真实用户表格、处理报告或登录凭据。性别属于个人信息，请只处理你有权限修改的记录。Chrome 会话需要手动登录，任务结束后会关闭。

## Limitations / 已知限制

The tool currently reads the first worksheet in `.xlsx` files and assumes one applicant match per UID. Sumsub UI changes may require locator updates. It was developed for an existing Sumsub workflow; it is not an official Sumsub client or SDK.

目前只读取 `.xlsx` 的第一个工作表，并要求每个 UID 只匹配一个申请人。Sumsub 界面改变时，可能需要更新定位规则。本项目不是 Sumsub 官方客户端或 SDK。
