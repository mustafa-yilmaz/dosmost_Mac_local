import React, { useState, useEffect } from 'react';
import { office365Service } from '../services/office365-service';
import { ExcelFile, Worksheet } from '../types/excel.types';

interface ExcelFilePickerProps {
  onSelect: (driveId: string, itemId: string, worksheetName: string) => void;
}

export const ExcelFilePicker: React.FC<ExcelFilePickerProps> = ({ onSelect }) => {
  const [files, setFiles] = useState<ExcelFile[]>([]);
  const [worksheets, setWorksheets] = useState<Worksheet[]>([]);
  const [selectedFile, setSelectedFile] = useState<ExcelFile | null>(null);
  const [driveId, setDriveId] = useState<string>('');
  const [isLoading, setIsLoading] = useState(false);

  useEffect(() => {
    loadFiles();
  }, []);

  const loadFiles = async () => {
    setIsLoading(true);
    try {
      const drive = await office365Service.getUserDrive();
      setDriveId(drive.id);
      const filesList = await office365Service.listExcelFiles();
      setFiles(filesList);
    } catch (error) {
      console.error('Error loading files:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleFileSelect = async (file: ExcelFile) => {
    setSelectedFile(file);
    setIsLoading(true);
    try {
      const worksheetsList = await office365Service.getWorksheets(driveId, file.id);
      setWorksheets(worksheetsList);
    } catch (error) {
      console.error('Error loading worksheets:', error);
    } finally {
      setIsLoading(false);
    }
  };

  const handleWorksheetSelect = (worksheet: Worksheet) => {
    if (selectedFile) {
      onSelect(driveId, selectedFile.id, worksheet.name);
    }
  };

  return (
    <div className="space-y-4">
      <div>
        <h3 className="text-lg font-semibold mb-2">Select Excel File</h3>
        {isLoading && <p className="text-sm text-gray-500">Loading...</p>}
        <div className="border rounded-lg divide-y max-h-60 overflow-y-auto">
          {files.map((file) => (
            <div
              key={file.id}
              onClick={() => handleFileSelect(file)}
              className={`p-3 cursor-pointer hover:bg-gray-50 ${
                selectedFile?.id === file.id ? 'bg-blue-50 border-blue-500' : ''
              }`}
            >
              <div className="font-medium">{file.name}</div>
              <div className="text-xs text-gray-500">
                Last modified: {new Date(file.lastModifiedDateTime).toLocaleDateString()}
              </div>
            </div>
          ))}
          {files.length === 0 && !isLoading && (
            <div className="p-4 text-center text-gray-500">No Excel files found</div>
          )}
        </div>
      </div>

      {worksheets.length > 0 && (
        <div>
          <h3 className="text-lg font-semibold mb-2">Select Worksheet</h3>
          <div className="border rounded-lg divide-y">
            {worksheets.map((worksheet) => (
              <div
                key={worksheet.id}
                onClick={() => handleWorksheetSelect(worksheet)}
                className="p-3 cursor-pointer hover:bg-gray-50"
              >
                <div className="font-medium">{worksheet.name}</div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};
