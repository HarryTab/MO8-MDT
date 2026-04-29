# GitHub Pages Hosting

This project includes a GitHub Actions workflow at:

```text
.github/workflows/pages.yml
```

It publishes the `frontend/` folder to GitHub Pages whenever you push to the `main` branch.

## Steps

1. Create a new GitHub repository.
2. Upload the whole project folder, not just `frontend/`.
3. Go to the repository on GitHub.
4. Open `Settings > Pages`.
5. Under `Build and deployment`, set `Source` to `GitHub Actions`.
6. Go to the `Actions` tab.
7. Run the `Deploy frontend to GitHub Pages` workflow, or push a change to `main`.
8. Wait for the workflow to finish.
9. Open the Pages URL shown in `Settings > Pages`.

The final URL will usually look like:

```text
https://YOUR_USERNAME.github.io/YOUR_REPOSITORY_NAME/
```

## If You Upload Through the GitHub Website

Make sure these folders/files are included:

```text
.github/workflows/pages.yml
apps-script/Code.gs
docs/
frontend/index.html
frontend/styles.css
frontend/app.js
README.md
```

The `.github` folder is important. Without it, GitHub will not know to publish the `frontend/` folder.

## If the Site Loads but Login Fails

Check:

- `frontend/app.js` has the correct Apps Script web app URL.
- Apps Script deployment is set to `Who has access: Anyone`.
- You created the first admin user in the Sheet.
- The Apps Script deployment is updated after any backend code changes.
