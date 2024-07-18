import React, { useState, useEffect } from 'react';

const BSAProcessor: React.FC = () => {
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const startBSAProcess = async () => {
    setIsProcessing(true);
    setProgress(0);
    setDownloadUrl(null);
    setError(null);

    const eventSource = new EventSource('http://localhost:3000/run-bsa?folder=LANDCRAFT-RECREATIONS');

    eventSource.onmessage = (event) => {
      const data = event.data;
      console.log('Received event:', data);

      if (data.startsWith("BSA process started")) {
        setProgress(10);
      } else if (data.startsWith("BSA process completed")) {
        setProgress(90);
      } else if (data.startsWith("Download URL:")) {
        const url = data.split(': ')[1];
        setDownloadUrl(url);
        setProgress(100);
        setIsProcessing(false);
        eventSource.close();
      } else if (data.startsWith("Error:")) {
        setError(data);
        setIsProcessing(false);
        eventSource.close();
      }
    };

    eventSource.onerror = (error) => {
      console.error('EventSource failed:', error);
      setError(`An error occurred during the BSA process: ${JSON.stringify(error)}`);
      setIsProcessing(false);
      eventSource.close();
    };
  };

  return (
    <div className="bsa-processor">
      <h1>BSA Processor</h1>

      {!isProcessing && !downloadUrl && (
        <button onClick={startBSAProcess} disabled={isProcessing}>
          Start BSA Process
        </button>
      )}

      {isProcessing && (
        <div className="progress-bar">
          <div className="progress" style={{ width: `${progress}%` }}></div>
        </div>
      )}

      {downloadUrl && (
        <a href={downloadUrl} download className="download-button">
          Download Result
        </a>
      )}

      {error && <p className="error">{error}</p>}
    </div>
  );
};

export default BSAProcessor;
