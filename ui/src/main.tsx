import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { TooltipProvider } from '@/components/ui/tooltip';
import App from '@/App';
import '@/styles.css';

const container = document.getElementById('root');
if (!container) throw new Error('#root is missing from index.html');

createRoot(container).render(
  <StrictMode>
    <TooltipProvider delayDuration={250} skipDelayDuration={300}>
      <App />
    </TooltipProvider>
  </StrictMode>
);
