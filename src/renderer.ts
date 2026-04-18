import './index.css';

const app = document.querySelector<HTMLDivElement>('#app');

if (!app) {
  throw new Error('App root was not found');
}

const shell = document.createElement('main');
shell.className = 'shell';

const eyebrow = document.createElement('p');
eyebrow.className = 'shell__eyebrow';
eyebrow.textContent = window.appShell.name;

const title = document.createElement('h1');
title.className = 'shell__title';
title.textContent = 'Hello world';

const body = document.createElement('p');
body.className = 'shell__body';
body.textContent =
  'Electron Forge, Vite, and TypeScript are wired up. This is the minimal app shell.';

shell.append(eyebrow, title, body);
app.append(shell);
