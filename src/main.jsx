import { StrictMode, Component } from 'react';
import { createRoot } from 'react-dom/client';
import './index.css';
import App from './App.jsx';

/** React 渲染鏈錯誤時預設整頁空白；改為可讀錯誤與重新載入 */
class RootErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  componentDidCatch(error, info) {
    console.error('route-planner: UI render error', error, info?.componentStack);
  }

  render() {
    if (this.state.error) {
      const msg = this.state.error?.message ?? String(this.state.error);
      return (
        <div
          style={{
            padding: '24px',
            fontFamily: 'system-ui, sans-serif',
            maxWidth: 560,
            margin: '0 auto',
            color: '#0f172a',
            boxSizing: 'border-box',
          }}
        >
          <h1 style={{ fontSize: '18px', marginTop: 0 }}>Route Planner hit an error</h1>
          <p style={{ color: '#475569', fontSize: '14px', lineHeight: 1.5 }}>
            Without this screen the window would stay blank. Try reloading. If it repeats, clear saved import from the sidebar,
            or reinstall the desktop app.
          </p>
          <pre
            style={{
              fontSize: '12px',
              background: '#f1f5f9',
              padding: '12px',
              borderRadius: '8px',
              overflow: 'auto',
              whiteSpace: 'pre-wrap',
              wordBreak: 'break-word',
            }}
          >
            {msg}
          </pre>
          <button
            type="button"
            onClick={() => window.location.reload()}
            style={{
              marginTop: '16px',
              padding: '10px 18px',
              fontWeight: 700,
              cursor: 'pointer',
              borderRadius: '8px',
              border: '1px solid #cbd5e1',
              background: '#fff',
            }}
          >
            Reload
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

const rootEl = document.getElementById('root');
if (!rootEl) {
  document.body.innerHTML =
    '<p style="font-family:system-ui;padding:16px;margin:0">Missing #root — invalid index.html.</p>';
} else {
  window.addEventListener('error', (ev) => {
    console.error('route-planner: window error', ev.error ?? ev.message);
  });
  window.addEventListener('unhandledrejection', (ev) => {
    console.error('route-planner: unhandled rejection', ev.reason);
  });

  createRoot(rootEl).render(
    <StrictMode>
      <RootErrorBoundary>
        <App />
      </RootErrorBoundary>
    </StrictMode>,
  );
}
