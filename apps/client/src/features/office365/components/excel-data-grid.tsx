import React, { useState } from 'react';
import { ExcelData, RowData } from '../types/excel.types';

interface ExcelDataGridProps {
  data: ExcelData;
  onCreateRow: (rowData: RowData) => Promise<void>;
  onUpdateRow: (rowIndex: string, rowData: RowData) => Promise<void>;
  onDeleteRow: (rowIndex: string) => Promise<void>;
}

export const ExcelDataGrid: React.FC<ExcelDataGridProps> = ({
  data,
  onCreateRow,
  onUpdateRow,
  onDeleteRow,
}) => {
  const [editingRow, setEditingRow] = useState<number | null>(null);
  const [editData, setEditData] = useState<RowData>({});
  const [isAddingRow, setIsAddingRow] = useState(false);
  const [newRowData, setNewRowData] = useState<RowData>({});

  if (!data || !data.values || data.values.length === 0) {
    return <div className="p-4 text-center text-gray-500">No data available</div>;
  }

  const headers = data.values[0] || [];
  const rows = data.values.slice(1);

  const handleEdit = (rowIndex: number) => {
    setEditingRow(rowIndex);
    const rowData: RowData = {};
    headers.forEach((header, colIndex) => {
      rowData[header] = rows[rowIndex][colIndex];
    });
    setEditData(rowData);
  };

  const handleSave = async (rowIndex: number) => {
    try {
      // Row index in Excel is 1-based and we need to account for the header
      await onUpdateRow(String(rowIndex + 2), editData);
      setEditingRow(null);
      setEditData({});
    } catch (error) {
      console.error('Error saving row:', error);
    }
  };

  const handleDelete = async (rowIndex: number) => {
    if (confirm('Are you sure you want to delete this row?')) {
      try {
        await onDeleteRow(String(rowIndex + 2));
      } catch (error) {
        console.error('Error deleting row:', error);
      }
    }
  };

  const handleAddRow = async () => {
    try {
      await onCreateRow(newRowData);
      setIsAddingRow(false);
      setNewRowData({});
    } catch (error) {
      console.error('Error adding row:', error);
    }
  };

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center">
        <h3 className="text-lg font-semibold">
          Excel Data ({rows.length} rows)
        </h3>
        <button
          onClick={() => setIsAddingRow(true)}
          className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700"
        >
          Add Row
        </button>
      </div>

      <div className="overflow-x-auto border rounded-lg">
        <table className="min-w-full divide-y divide-gray-200">
          <thead className="bg-gray-50">
            <tr>
              {headers.map((header, index) => (
                <th
                  key={index}
                  className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider"
                >
                  {header}
                </th>
              ))}
              <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                Actions
              </th>
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-200">
            {isAddingRow && (
              <tr className="bg-yellow-50">
                {headers.map((header, colIndex) => (
                  <td key={colIndex} className="px-4 py-2">
                    <input
                      type="text"
                      value={newRowData[header] || ''}
                      onChange={(e) =>
                        setNewRowData({ ...newRowData, [header]: e.target.value })
                      }
                      className="w-full px-2 py-1 border rounded"
                      placeholder={header}
                    />
                  </td>
                ))}
                <td className="px-4 py-2">
                  <div className="flex gap-2">
                    <button
                      onClick={handleAddRow}
                      className="px-3 py-1 bg-green-600 text-white rounded text-sm hover:bg-green-700"
                    >
                      Save
                    </button>
                    <button
                      onClick={() => {
                        setIsAddingRow(false);
                        setNewRowData({});
                      }}
                      className="px-3 py-1 bg-gray-600 text-white rounded text-sm hover:bg-gray-700"
                    >
                      Cancel
                    </button>
                  </div>
                </td>
              </tr>
            )}
            {rows.map((row, rowIndex) => (
              <tr key={rowIndex} className={editingRow === rowIndex ? 'bg-blue-50' : ''}>
                {headers.map((header, colIndex) => (
                  <td key={colIndex} className="px-4 py-2 text-sm">
                    {editingRow === rowIndex ? (
                      <input
                        type="text"
                        value={editData[header] || ''}
                        onChange={(e) =>
                          setEditData({ ...editData, [header]: e.target.value })
                        }
                        className="w-full px-2 py-1 border rounded"
                      />
                    ) : (
                      row[colIndex]
                    )}
                  </td>
                ))}
                <td className="px-4 py-2 text-sm">
                  {editingRow === rowIndex ? (
                    <div className="flex gap-2">
                      <button
                        onClick={() => handleSave(rowIndex)}
                        className="px-3 py-1 bg-green-600 text-white rounded text-sm hover:bg-green-700"
                      >
                        Save
                      </button>
                      <button
                        onClick={() => {
                          setEditingRow(null);
                          setEditData({});
                        }}
                        className="px-3 py-1 bg-gray-600 text-white rounded text-sm hover:bg-gray-700"
                      >
                        Cancel
                      </button>
                    </div>
                  ) : (
                    <div className="flex gap-2">
                      <button
                        onClick={() => handleEdit(rowIndex)}
                        className="px-3 py-1 bg-blue-600 text-white rounded text-sm hover:bg-blue-700"
                      >
                        Edit
                      </button>
                      <button
                        onClick={() => handleDelete(rowIndex)}
                        className="px-3 py-1 bg-red-600 text-white rounded text-sm hover:bg-red-700"
                      >
                        Delete
                      </button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
};
