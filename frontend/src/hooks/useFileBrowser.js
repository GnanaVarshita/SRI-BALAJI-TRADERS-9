/**
 * useFileBrowser Hook
 * Encapsulates thread-safe file and folder browsing states and callbacks.
 */

import { useState } from 'react';
import { api } from '../services/api';

export function useFileBrowser() {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  const browseFile = async (onSuccess) => {
    setLoading(true);
    setError(null);
    try {
      const data = await api.browseFile();
      if (data.success && data.filePath) {
        if (onSuccess) onSuccess(data.filePath);
        return data.filePath;
      } else if (data.message) {
        setError(data.message);
      }
    } catch (err) {
      console.error(err);
      setError('Failed to open file browser dialog.');
    } finally {
      setLoading(false);
    }
    return null;
  };

  const browseFolder = async (onSuccess) => {
    setLoading(true);
    setError(null);
    try {
      const data = await api.browseFolder();
      if (data.success && data.folderPath) {
        if (onSuccess) onSuccess(data.folderPath);
        return data.folderPath;
      } else if (data.message) {
        setError(data.message);
      }
    } catch (err) {
      console.error(err);
      setError('Failed to open folder browser dialog.');
    } finally {
      setLoading(false);
    }
    return null;
  };

  return {
    loading,
    error,
    setError,
    browseFile,
    browseFolder,
  };
}
