# 项目规则（project_rules）

## Git 操作约束

- **未经用户明确指令，禁止执行 `git commit` 与 `git push`。**
- 用户需通过明确命令（如"提交并推送""push 一下""commit 这次改动"）或在 AI 询问后明确确认，AI 才可执行提交/推送。
- `git status`、`git diff`、`git log`、`git add` 等只读/暂存类操作可按需执行，无需额外确认。
- 涉及改写历史或远程的危险操作（`reset --hard`、`rebase`、`push --force`、`branch -D` 等）一律需用户确认。

## 环境备注

- Git 可执行文件路径：`E:\Git\cmd\git.exe`（系统 PATH 中可能不可直接调用 `git`，必要时用全路径）。
- 远程仓库：`https://github.com/MuLiKun/zemax-skills.git`，主分支 `main`。
- 公差分析项目位于 `tolerance_analysis/`；Python 解释器使用工作区共享 venv：`.venv\Scripts\python.exe`（`.venv` 不入库）。
