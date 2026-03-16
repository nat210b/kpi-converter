import React, { useRef, useState } from 'react'

interface Props {
  onFile: (file: File) => void
  disabled?: boolean
  fileName?: string
  fileSize?: string
}

export default function DropZone({ onFile, disabled, fileName, fileSize }: Props) {
  const [dragging, setDragging] = useState(false)
  const inputRef = useRef<HTMLInputElement>(null)

  function handleDragOver(e: React.DragEvent) {
    e.preventDefault()
    if (!disabled) setDragging(true)
  }

  function handleDragLeave() { setDragging(false) }

  function handleDrop(e: React.DragEvent) {
    e.preventDefault()
    setDragging(false)
    if (disabled) return
    const file = e.dataTransfer.files[0]
    if (file) onFile(file)
  }

  function handleChange(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (file) onFile(file)
    e.target.value = ''
  }

  const classes = ['drop-zone', dragging ? 'dragging' : '', disabled ? 'disabled' : '']
    .filter(Boolean).join(' ')

  return (
    <div
      className={classes}
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      onClick={() => !disabled && inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: 'none' }}
        onChange={handleChange}
        disabled={disabled}
      />
      {fileName ? (
        <div className="file-info">
          <span className="file-icon">📄</span>
          <div>
            <div className="file-name">{fileName}</div>
            <div className="file-meta">{fileSize}</div>
          </div>
          <span className="change-hint">คลิกเพื่อเปลี่ยนไฟล์</span>
        </div>
      ) : (
        <>
          <div className="drop-icon">📂</div>
          <div className="drop-label">วาง หรือ คลิกเพื่อเลือกไฟล์</div>
          <div className="drop-sub">รองรับ .xlsx / .xls — ต้องมี sheet ชื่อ "KPI ปี 2568"</div>
        </>
      )}
    </div>
  )
}
