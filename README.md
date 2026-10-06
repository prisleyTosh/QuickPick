# QuickPick Admin

Vite + React version of the uploaded QuickPick `App (4).jsx`, prepared for GitHub Pages.

## Local setup

```bash
npm install
npm run dev
```

## Production build

```bash
npm run build
npm run preview
```

## GitHub Pages

The included GitHub Actions workflow builds and deploys the app whenever `main` is pushed.

If the repository is not named `quickpick`, update `base` in `vite.config.js` to:

```js
base: "/YOUR-REPOSITORY-NAME/"
```
