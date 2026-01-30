import { useCallback } from 'react'
import { useDropzone } from 'react-dropzone'

function Header({
  validatorType,
  onValidatorTypeChange,
  onFileUpload,
  onSaveFile,
  loading,
  fileName,
  results,
  sessionId,
}) {
  const onDrop = useCallback(
    (acceptedFiles) => {
      if (acceptedFiles.length > 0) {
        onFileUpload(acceptedFiles[0])
      }
    },
    [onFileUpload]
  )

  const { getRootProps, getInputProps, open } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
    },
    noClick: true,
    noKeyboard: true,
    multiple: false,
  })

  const passed = results.filter((r) => r.status === 'PASS').length
  const failed = results.length - passed

  return (
    <header className="bg-main-bg px-8 py-6 border-b border-gray-200" {...getRootProps()}>
      <input {...getInputProps()} />

      <div className="flex items-start justify-between">
        {/* Left Side - Title & Validator Type */}
        <div>
          {/* Validator Type Selection */}
          <div className="flex items-center gap-4 mb-4">
            <span className="text-sm font-semibold text-text-secondary">
              Validator Type:
            </span>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="radio"
                name="validatorType"
                value="QUOTE_WIN"
                checked={validatorType === 'QUOTE_WIN'}
                onChange={(e) => onValidatorTypeChange(e.target.value)}
                className="w-4 h-4 text-primary focus:ring-primary"
              />
              <span className="text-sm text-text-primary">Quote Win (4 Rules)</span>
            </label>
            <label className="flex items-center gap-2 cursor-pointer">
              <input
                type="radio"
                name="validatorType"
                value="BOM"
                checked={validatorType === 'BOM'}
                onChange={(e) => onValidatorTypeChange(e.target.value)}
                className="w-4 h-4 text-primary focus:ring-primary"
              />
              <span className="text-sm text-text-primary">BOM Matrix (18 Rules)</span>
            </label>
          </div>

          {/* Title */}
          <h1 className="text-3xl font-bold text-text-primary">
            {loading ? 'Validating...' : results.length > 0 ? 'Validation Complete' : 'Welcome to Validator'}
          </h1>

          {/* Subtitle */}
          <p className="text-sm text-text-secondary mt-1">
            {fileName
              ? `File: ${fileName}${results.length > 0 ? ` | ${passed} passed, ${failed} failed` : ''}`
              : `Upload an Excel file to validate ${validatorType === 'QUOTE_WIN' ? 'Quote Win' : 'BOM Matrix'} structure`}
          </p>
        </div>

        {/* Right Side - Buttons */}
        <div className="flex items-center gap-3">
          <button
            onClick={open}
            disabled={loading}
            className="px-6 py-3 bg-primary hover:bg-primary-hover text-white font-semibold rounded-lg transition-all-200 disabled:opacity-50 disabled:cursor-not-allowed"
          >
            Select File
          </button>

          {validatorType === 'QUOTE_WIN' && results.length > 0 && (
            <button
              onClick={onSaveFile}
              disabled={loading || !sessionId}
              className="px-6 py-3 bg-accent-pink hover:bg-pink-300 text-text-primary font-semibold rounded-lg transition-all-200 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Save File
            </button>
          )}
        </div>
      </div>

      {/* Loading Indicator */}
      {loading && (
        <div className="mt-4 flex items-center gap-2 text-primary">
          <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
            <circle
              className="opacity-25"
              cx="12"
              cy="12"
              r="10"
              stroke="currentColor"
              strokeWidth="4"
              fill="none"
            />
            <path
              className="opacity-75"
              fill="currentColor"
              d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
            />
          </svg>
          <span className="text-sm font-medium">Processing Excel file...</span>
        </div>
      )}
    </header>
  )
}

export default Header
