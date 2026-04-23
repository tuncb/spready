import '@glideapps/glide-data-grid/dist/index.css';
import './index.css';

import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';

import App from './App';
import { ChartEditorWindow } from './ChartEditorWindow';

const app = document.querySelector<HTMLDivElement>('#app');

if (!app) {
  throw new Error('App root was not found');
}

createRoot(app).render(
  <StrictMode>
    {new URLSearchParams(window.location.search).get('view') === 'chart-editor' ? (
      <ChartEditorWindow />
    ) : (
      <App />
    )}
  </StrictMode>,
);
