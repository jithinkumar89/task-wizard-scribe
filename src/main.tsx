
import { createRoot } from 'react-dom/client'
import App from './App.tsx'
import './index.css'

// Set application version
window.APP_VERSION = "1.0";

createRoot(document.getElementById("root")!).render(<App />);
