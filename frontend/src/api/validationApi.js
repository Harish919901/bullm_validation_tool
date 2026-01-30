import axios from 'axios'

const API_BASE = '/api'

/**
 * Upload file and run validation
 * @param {File} file - Excel file to validate
 * @param {string} validatorType - 'QUOTE_WIN' or 'BOM'
 * @returns {Promise<Object>} Validation response
 */
export const uploadAndValidate = async (file, validatorType) => {
  const formData = new FormData()
  formData.append('file', file)
  formData.append('validator_type', validatorType)

  const response = await axios.post(`${API_BASE}/validate`, formData, {
    headers: {
      'Content-Type': 'multipart/form-data',
    },
  })

  return response.data
}

/**
 * Get rules information for a validator type
 * @param {string} validatorType - 'QUOTE_WIN' or 'BOM'
 * @returns {Promise<Object>} Rules response
 */
export const getRules = async (validatorType) => {
  const response = await axios.get(`${API_BASE}/rules/${validatorType}`)
  return response.data
}

/**
 * Export results to CSV
 * @param {string} sessionId - Session ID from validation
 * @returns {Promise<Blob>} CSV file blob
 */
export const exportCSV = async (sessionId) => {
  const formData = new FormData()
  formData.append('session_id', sessionId)

  const response = await axios.post(`${API_BASE}/export/csv`, formData, {
    responseType: 'blob',
  })

  // Trigger download
  const url = window.URL.createObjectURL(new Blob([response.data]))
  const link = document.createElement('a')
  link.href = url
  link.setAttribute('download', 'validation_report.csv')
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.URL.revokeObjectURL(url)
}

/**
 * Export results to Excel
 * @param {string} sessionId - Session ID from validation
 * @returns {Promise<Blob>} Excel file blob
 */
export const exportExcel = async (sessionId) => {
  const formData = new FormData()
  formData.append('session_id', sessionId)

  const response = await axios.post(`${API_BASE}/export/excel`, formData, {
    responseType: 'blob',
  })

  // Trigger download
  const url = window.URL.createObjectURL(new Blob([response.data]))
  const link = document.createElement('a')
  link.href = url
  link.setAttribute('download', 'validation_report.xlsx')
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.URL.revokeObjectURL(url)
}

/**
 * Save validated file (Quote Win only)
 * @param {string} sessionId - Session ID from validation
 * @returns {Promise<Blob>} Validated Excel file blob
 */
export const saveValidatedFile = async (sessionId) => {
  const formData = new FormData()
  formData.append('session_id', sessionId)

  const response = await axios.post(`${API_BASE}/save`, formData, {
    responseType: 'blob',
  })

  // Trigger download
  const url = window.URL.createObjectURL(new Blob([response.data]))
  const link = document.createElement('a')
  link.href = url
  link.setAttribute('download', 'validated_file.xlsx')
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.URL.revokeObjectURL(url)
}

/**
 * Delete session and cleanup
 * @param {string} sessionId - Session ID to delete
 */
export const deleteSession = async (sessionId) => {
  await axios.delete(`${API_BASE}/session/${sessionId}`)
}
