
import { createRoot } from 'react-dom/client'
import App from './App.tsx'
import './index.css'

// Declare APP_VERSION property on Window interface
declare global {
  interface Window {
    APP_VERSION: string;
  }
}

// Set application version
window.APP_VERSION = "1.0";

createRoot(document.getElementById("root")!).render(<App />);
