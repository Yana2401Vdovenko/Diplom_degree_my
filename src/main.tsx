import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.tsx';
import './index.css';
import './styles/ComfortLayout.css';

// Якщо у папці src є файли глобальних стилів (наприклад, index.css), 
// і вони тобі потрібні, залиш цей рядок. Якщо вони псують вигляд, його можна видалити.
// import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
