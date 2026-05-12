# Agent Workflow Guidelines

## Git commit and push

- Stage all modified and untracked files that are **not excluded by `.gitignore`**.
- Commit with a concise message, then push to `origin main`.
- Do **not** stage files listed in `.gitignore`.

## Release management

When `xlsm_devkit.bas` has changed since the last release, create a new GitHub release:

1. **Determine the version** using Semantic Versioning based on the `[Unreleased]` section of `CHANGELOG.md`:
   - New feature added → bump MINOR (e.g. `1.0.x` → `1.1.0`)
   - Bug fix or non-breaking change only → bump PATCH (e.g. `1.1.0` → `1.1.1`)
   - Breaking change → bump MAJOR (e.g. `1.1.0` → `2.0.0`)

2. **Update `CHANGELOG.md`**: replace `## [Unreleased]` with a new `## [Unreleased]` section (empty) followed by `## [x.y.z] - YYYY-MM-DD`, then commit and push.

3. **Tag and release**:
   ```
   git tag vX.Y.Z
   git push origin main
   git push origin vX.Y.Z
   gh release create vX.Y.Z "xlsm_devkit.bas#xlsm_devkit.bas" \
     --title "vX.Y.Z" \
     --notes "<content of the new CHANGELOG section>"
   ```

4. **Always attach `xlsm_devkit.bas`** as a release asset so users can download it directly without cloning the repository.
