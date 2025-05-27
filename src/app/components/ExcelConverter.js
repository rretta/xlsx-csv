'use client'

import { useState, useCallback } from 'react'
import { useDropzone } from 'react-dropzone'
import * as XLSX from 'xlsx'

export default function ExcelConverter() {
  const [excelData, setExcelData] = useState(null)
  const [csvData, setCsvData] = useState(null)
  const [error, setError] = useState(null)
  const [success, setSuccess] = useState(null)

  const onDrop = useCallback((acceptedFiles) => {
    const file = acceptedFiles[0]
    if (!file) return

    if (!file.name.endsWith('.xlsx')) {
      setError('Por favor, sube un archivo .xlsx')
      return
    }

    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        const firstSheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[firstSheetName]
        
        // Convert to JSON
        const jsonData = XLSX.utils.sheet_to_json(worksheet)
        setExcelData(jsonData)

        // Convert to CSV
        const csv = XLSX.utils.sheet_to_csv(worksheet)
        setCsvData(csv)

        setSuccess('El archivo se ha convertido exitosamente')
        setError(null)
      } catch (error) {
        setError('Error al procesar el archivo')
        setSuccess(null)
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx']
    },
    multiple: false
  })

  const downloadCSV = () => {
    if (!csvData) return

    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' })
    const link = document.createElement('a')
    const url = URL.createObjectURL(blob)
    
    link.setAttribute('href', url)
    link.setAttribute('download', 'converted.csv')
    link.style.visibility = 'hidden'
    
    document.body.appendChild(link)
    link.click()
    document.body.removeChild(link)
  }

  return (
    <div className="container mx-auto px-4 py-10">
      <div className="flex flex-col items-center space-y-8">
        <h1 className="text-3xl font-bold">Convertidor de Excel a CSV</h1>
        
        <div
          {...getRootProps()}
          className={`p-10 border-2 border-dashed rounded-lg w-full text-center cursor-pointer transition-colors
            ${isDragActive ? 'border-blue-500 bg-blue-50' : 'border-gray-200 hover:border-blue-500'}`}
        >
          <input {...getInputProps()} />
          {isDragActive ? (
            <p className="text-gray-600">Suelta el archivo aquí...</p>
          ) : (
            <p className="text-gray-600">Arrastra y suelta un archivo .xlsx aquí, o haz clic para seleccionar</p>
          )}
        </div>

        {error && (
          <div className="w-full p-4 bg-red-100 border border-red-400 text-red-700 rounded">
            {error}
          </div>
        )}

        {success && (
          <div className="w-full p-4 bg-green-100 border border-green-400 text-green-700 rounded">
            {success}
          </div>
        )}

        {excelData && (
          <div className="w-full space-y-4">
            <p className="font-bold">Vista previa de los datos:</p>
            <div className="p-4 border border-gray-200 rounded-md w-full max-h-[300px] overflow-y-auto">
              <pre className="whitespace-pre-wrap">{JSON.stringify(excelData.slice(0, 5), null, 2)}</pre>
            </div>
            
            <button
              onClick={downloadCSV}
              disabled={!csvData}
              className={`px-4 py-2 rounded-md text-white font-medium
                ${csvData ? 'bg-blue-500 hover:bg-blue-600' : 'bg-gray-400 cursor-not-allowed'}`}
            >
              Descargar CSV
            </button>
          </div>
        )}
      </div>
    </div>
  )
} 