import { contextBridge } from 'electron';

contextBridge.exposeInMainWorld('appShell', {
  name: 'Spready',
});
