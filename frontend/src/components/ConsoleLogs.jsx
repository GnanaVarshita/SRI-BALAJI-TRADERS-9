import React from 'react';

function ConsoleLogs({ logs, consoleRef }) {
  return (
    <section className="card full-width">
      <h2>Console Logs Output</h2>
      <div className="console" ref={consoleRef}>
        {logs.length === 0 ? (
          <div className="console-line system">Standby. Ready to download attachments...</div>
        ) : (
          logs.map((line, idx) => (
            <div key={idx} className="console-line">
              {line}
            </div>
          ))
        )}
      </div>
    </section>
  );
}

export default ConsoleLogs;
