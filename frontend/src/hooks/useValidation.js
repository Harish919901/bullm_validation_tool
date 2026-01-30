import { useState, useCallback } from 'react'
import * as api from '../api/validationApi'

/**
 * Custom hook for validation operations
 */
export const useValidation = () => {
  const [results, setResults] = useState([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)
  const [fileName, setFileName] = useState(null)
  const [sessionId, setSessionId] = useState(null)

  /**
   * Upload file and run validation
   */
  const uploadAndValidate = useCallback(async (file, validatorType) => {
    setLoading(true)
    setError(null)

    try {
      const response = await api.uploadAndValidate(file, validatorType)
      setResults(response.results)
      setFileName(response.file_name)
      setSessionId(response.session_id)
    } catch (err) {
      setError(err.response?.data?.detail || err.message || 'Validation failed')
      setResults([])
    } finally {
      setLoading(false)
    }
  }, [])

  /**
   * Export results to CSV
   */
  const exportCSV = useCallback(async () => {
    if (!sessionId) {
      setError('No validation results to export')
      return
    }

    try {
      await api.exportCSV(sessionId)
    } catch (err) {
      setError(err.response?.data?.detail || err.message || 'Export failed')
    }
  }, [sessionId])

  /**
   * Export results to Excel
   */
  const exportExcel = useCallback(async () => {
    if (!sessionId) {
      setError('No validation results to export')
      return
    }

    try {
      await api.exportExcel(sessionId)
    } catch (err) {
      setError(err.response?.data?.detail || err.message || 'Export failed')
    }
  }, [sessionId])

  /**
   * Save validated file (Quote Win only)
   */
  const saveValidatedFile = useCallback(async () => {
    if (!sessionId) {
      setError('No validation results to save')
      return
    }

    try {
      await api.saveValidatedFile(sessionId)
    } catch (err) {
      setError(err.response?.data?.detail || err.message || 'Save failed')
    }
  }, [sessionId])

  /**
   * Clear all results
   */
  const clearResults = useCallback(() => {
    setResults([])
    setFileName(null)
    setError(null)

    // Cleanup session on server
    if (sessionId) {
      api.deleteSession(sessionId).catch(() => {})
      setSessionId(null)
    }
  }, [sessionId])

  return {
    results,
    loading,
    error,
    fileName,
    sessionId,
    uploadAndValidate,
    exportCSV,
    exportExcel,
    saveValidatedFile,
    clearResults,
  }
}

export default useValidation
