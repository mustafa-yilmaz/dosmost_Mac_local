import { useState, useEffect } from 'react';
import { office365Service } from '../services/office365-service';
import { ExcelData, RowData } from '../types/excel.types';

export const useExcelData = (
  driveId: string | null,
  itemId: string | null,
  worksheetName: string | null,
) => {
  const [data, setData] = useState<ExcelData | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const fetchData = async () => {
    if (!driveId || !itemId || !worksheetName) return;

    setIsLoading(true);
    setError(null);

    try {
      const excelData = await office365Service.getExcelData(
        driveId,
        itemId,
        worksheetName,
      );
      setData(excelData);
    } catch (err: any) {
      setError(err.message || 'Failed to fetch Excel data');
      console.error('Error fetching Excel data:', err);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, [driveId, itemId, worksheetName]);

  const createRow = async (rowData: RowData) => {
    if (!driveId || !itemId || !worksheetName) return;

    try {
      await office365Service.createRow({
        driveId,
        itemId,
        worksheetName,
        data: rowData,
      });
      await fetchData(); // Refresh data
    } catch (err: any) {
      setError(err.message || 'Failed to create row');
      throw err;
    }
  };

  const updateRow = async (rowIndex: string, rowData: RowData) => {
    if (!driveId || !itemId || !worksheetName) return;

    try {
      await office365Service.updateRow({
        driveId,
        itemId,
        worksheetName,
        rowIndex,
        data: rowData,
      });
      await fetchData(); // Refresh data
    } catch (err: any) {
      setError(err.message || 'Failed to update row');
      throw err;
    }
  };

  const deleteRow = async (rowIndex: string) => {
    if (!driveId || !itemId || !worksheetName) return;

    try {
      await office365Service.deleteRow({
        driveId,
        itemId,
        worksheetName,
        rowIndex,
      });
      await fetchData(); // Refresh data
    } catch (err: any) {
      setError(err.message || 'Failed to delete row');
      throw err;
    }
  };

  return {
    data,
    isLoading,
    error,
    refresh: fetchData,
    createRow,
    updateRow,
    deleteRow,
  };
};
