import React, { useState } from 'react';
import { useMicrosoftAuth } from '../hooks/use-microsoft-auth';
import { useExcelData } from '../hooks/use-excel-data';
import { ExcelFilePicker } from './excel-file-picker';
import { ExcelDataGrid } from './excel-data-grid';

export const ExcelCRUD: React.FC = () => {
  const { account, isAuthenticated, isLoading, login, logout } = useMicrosoftAuth();
  const [driveId, setDriveId] = useState<string | null>(null);
  const [itemId, setItemId] = useState<string | null>(null);
  const [worksheetName, setWorksheetName] = useState<string | null>(null);

  const {
    data,
    isLoading: isDataLoading,
    error,
    createRow,
    updateRow,
    deleteRow,
    refresh,
  } = useExcelData(driveId, itemId, worksheetName);

  const handleFileSelect = (
    selectedDriveId: string,
    selectedItemId: string,
    selectedWorksheetName: string,
  ) => {
    setDriveId(selectedDriveId);
    setItemId(selectedItemId);
    setWorksheetName(selectedWorksheetName);
  };

  if (isLoading) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Loading...</p>
        </div>
      </div>
    );
  }

  if (!isAuthenticated) {
    return (
      <div className="flex items-center justify-center min-h-screen bg-gray-50">
        <div className="max-w-md w-full bg-white rounded-lg shadow-lg p-8 text-center">
          <div className="mb-6">
            <svg
              className="w-16 h-16 mx-auto text-blue-600"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth={2}
                d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
          </div>
          <h1 className="text-2xl font-bold text-gray-900 mb-2">
            Office 365 Excel CRUD
          </h1>
          <p className="text-gray-600 mb-6">
            Connect your Microsoft 365 account to manage Excel files online
          </p>
          <button
            onClick={login}
            className="w-full px-6 py-3 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 transition-colors"
          >
            Sign in with Microsoft
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="bg-white shadow">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
          <div className="flex justify-between items-center">
            <div>
              <h1 className="text-2xl font-bold text-gray-900">
                Office 365 Excel Manager
              </h1>
              {account && (
                <p className="text-sm text-gray-600">
                  Signed in as: {account.username}
                </p>
              )}
            </div>
            <button
              onClick={logout}
              className="px-4 py-2 bg-gray-600 text-white rounded hover:bg-gray-700"
            >
              Sign Out
            </button>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          <div className="lg:col-span-1">
            <div className="bg-white rounded-lg shadow p-6">
              <ExcelFilePicker onSelect={handleFileSelect} />
            </div>
          </div>

          <div className="lg:col-span-2">
            <div className="bg-white rounded-lg shadow p-6">
              {!driveId || !itemId || !worksheetName ? (
                <div className="text-center py-12">
                  <svg
                    className="w-16 h-16 mx-auto text-gray-400 mb-4"
                    fill="none"
                    stroke="currentColor"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                    />
                  </svg>
                  <h3 className="text-lg font-medium text-gray-900 mb-2">
                    No file selected
                  </h3>
                  <p className="text-gray-600">
                    Select an Excel file and worksheet to get started
                  </p>
                </div>
              ) : isDataLoading ? (
                <div className="text-center py-12">
                  <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
                  <p className="text-gray-600">Loading Excel data...</p>
                </div>
              ) : error ? (
                <div className="text-center py-12">
                  <div className="text-red-600 mb-4">
                    <svg
                      className="w-16 h-16 mx-auto"
                      fill="none"
                      stroke="currentColor"
                      viewBox="0 0 24 24"
                    >
                      <path
                        strokeLinecap="round"
                        strokeLinejoin="round"
                        strokeWidth={2}
                        d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                      />
                    </svg>
                  </div>
                  <h3 className="text-lg font-medium text-gray-900 mb-2">Error</h3>
                  <p className="text-red-600">{error}</p>
                  <button
                    onClick={refresh}
                    className="mt-4 px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700"
                  >
                    Retry
                  </button>
                </div>
              ) : data ? (
                <ExcelDataGrid
                  data={data}
                  onCreateRow={createRow}
                  onUpdateRow={updateRow}
                  onDeleteRow={deleteRow}
                />
              ) : null}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
