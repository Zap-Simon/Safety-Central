import { useState, useEffect, useMemo } from "react";
import { useQuery } from "@tanstack/react-query";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { Badge } from "@/components/ui/badge";
import { Skeleton } from "@/components/ui/skeleton";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Dialog, DialogContent, DialogTrigger, DialogTitle, DialogDescription } from "@/components/ui/dialog";
import { Pagination, PaginationContent, PaginationItem, PaginationNext, PaginationPrevious } from "@/components/ui/pagination";
import { RefreshCw, FileSpreadsheet, Download, AlertCircle, ImageOff, Search, ExternalLink, Check, X } from "lucide-react";
import { authService } from "@/auth/authService";
import type { SharePointExcelFile, WorkbookData, WorksheetData } from "@shared/schema";

interface ExcelTableProps {
  /** SharePoint site name (e.g., "HealthSafetyAdministration") */
  siteName: string;
  /** Folder path within the site (optional, defaults to root) */
  folderPath?: string;
  /** Title for the Excel table section */
  title?: string;
  /** Description for the Excel table section */
  description?: string;
  /** Number of rows per page for pagination */
  pageSize?: number;
}

// ImageCell component for displaying images from URLs
function ImageCell({ url, altText, columnName }: { url: string; altText?: string; columnName?: string }) {
  const [imageError, setImageError] = useState(false);
  const [imageLoaded, setImageLoaded] = useState(false);
  const [isDialogOpen, setIsDialogOpen] = useState(false);

  // Check if URL needs proxy (SharePoint URLs) or can be used directly
  const needsProxy = url.includes('sharepoint.com') || url.includes('graph.microsoft.com');
  const [imageUrl, setImageUrl] = useState<string>(needsProxy ? '' : url);

  // Get authenticated image URL for SharePoint images only
  useEffect(() => {
    if (needsProxy) {
      const fetchImageUrl = async () => {
        try {
          const token = await authService.getAccessToken();
          const response = await fetch(`/api/images/proxy?url=${encodeURIComponent(url)}`, {
            headers: {
              'Authorization': `Bearer ${token}`,
              'Accept': 'image/*'
            },
            credentials: 'include'
          });
          
          if (response.ok) {
            const blob = await response.blob();
            const objectUrl = URL.createObjectURL(blob);
            setImageUrl(objectUrl);
          } else {
            console.error('Image proxy failed:', response.status, response.statusText);
            setImageError(true);
            setImageLoaded(true);
          }
        } catch (error) {
          console.error('Error fetching image:', error);
          setImageError(true);
          setImageLoaded(true);
        }
      };

      fetchImageUrl();
    }

    // Cleanup function to revoke object URL
    return () => {
      if (needsProxy && imageUrl && imageUrl.startsWith('blob:')) {
        URL.revokeObjectURL(imageUrl);
      }
    };
  }, [url, needsProxy]);

  const handleImageLoad = () => {
    setImageLoaded(true);
    setImageError(false);
  };

  const handleImageError = () => {
    setImageError(true);
    setImageLoaded(true);
  };

  if (imageError) {
    return (
      <div className="flex items-center gap-1 text-gray-400" data-testid="image-error">
        <ImageOff className="h-4 w-4" />
        <span className="text-xs">Image unavailable</span>
      </div>
    );
  }

  return (
    <Dialog open={isDialogOpen} onOpenChange={setIsDialogOpen}>
      <DialogTrigger asChild>
        <div className="relative cursor-pointer group" data-testid="image-thumbnail">
          {!imageLoaded && (
            <Skeleton className="w-16 h-16 rounded" data-testid="image-skeleton" />
          )}
          <img
            src={imageUrl}
            alt={altText || "Equipment image"}
            className={`w-16 h-16 object-cover rounded border border-gray-200 transition-opacity group-hover:opacity-80 ${
              imageLoaded ? 'opacity-100' : 'opacity-0'
            }`}
            loading="lazy"
            onLoad={handleImageLoad}
            onError={handleImageError}
            data-testid="image-thumb"
          />
        </div>
      </DialogTrigger>
      <DialogContent className="max-w-[90vw] max-h-[90vh] p-4">
        <DialogTitle className="text-lg font-semibold mb-3">
          {columnName || 'Equipment Photo'}
        </DialogTitle>
        <DialogDescription className="sr-only">
          Full size view of equipment image
        </DialogDescription>
        <div className="flex items-center justify-center">
          <img
            src={imageUrl}
            alt={altText || "Equipment image"}
            className="max-w-full max-h-[80vh] object-contain rounded"
            data-testid="image-full"
          />
        </div>
        <div className="mt-3 text-sm text-gray-600 text-center space-y-2">
          <p>{altText || "Equipment image"}</p>
          <Button 
            variant="outline" 
            size="sm" 
            onClick={() => window.open(url, '_blank', 'noopener,noreferrer')}
            className="text-xs"
            data-testid="button-open-url"
          >
            <ExternalLink className="h-3 w-3 mr-1" />
            Open URL
          </Button>
        </div>
      </DialogContent>
    </Dialog>
  );
}

// Function to detect and format checkbox values
function formatCheckboxValue(value: any): { isCheckbox: boolean; checked: boolean } {
  if (value === null || value === undefined) {
    return { isCheckbox: false, checked: false };
  }
  
  const strValue = String(value).toLowerCase().trim();
  
  // Check for boolean values
  if (typeof value === 'boolean') {
    return { isCheckbox: true, checked: value };
  }
  
  // Check for numeric boolean (1/0)
  if (value === 1 || value === 0) {
    return { isCheckbox: true, checked: value === 1 };
  }
  
  // Check for text boolean values
  if (strValue === 'true' || strValue === 'yes' || strValue === 'checked' || strValue === 'tagged') {
    return { isCheckbox: true, checked: true };
  }
  
  if (strValue === 'false' || strValue === 'no' || strValue === 'unchecked' || strValue === 'untagged') {
    return { isCheckbox: true, checked: false };
  }
  
  // Check for checkmark symbols
  if (strValue === '✓' || strValue === '✔' || strValue === '☑' || strValue === '✅') {
    return { isCheckbox: true, checked: true };
  }
  
  if (strValue === '✗' || strValue === '✘' || strValue === '❌' || strValue === '☐') {
    return { isCheckbox: true, checked: false };
  }
  
  return { isCheckbox: false, checked: false };
}

// Function to detect and format dates
function formatCellValue(value: any, header: string): string {
  if (value === null || value === undefined) {
    return '';
  }
  
  const strValue = String(value);
  
  // Check if this might be a date column (more specific matching)
  const headerLower = header.toLowerCase();
  const isDateColumn = header && (
    headerLower.includes('date') ||
    headerLower.includes('last checked') ||
    headerLower.includes('next check') ||
    headerLower.includes('due') ||
    headerLower.includes('expiry') ||
    headerLower.includes('expire')
  );
  
  if (isDateColumn) {
    // Try to parse as Excel date (numeric) first - broader range
    const numValue = Number(value);
    if (!isNaN(numValue) && numValue > 20000 && numValue < 80000) {
      // Excel date serial number (broader range for older/future dates)
      const excelDate = new Date((numValue - 25569) * 86400 * 1000);
      return excelDate.toLocaleDateString();
    }
    
    // Try to parse as ISO date string
    const dateValue = new Date(strValue);
    if (!isNaN(dateValue.getTime()) && strValue.includes('-')) {
      return dateValue.toLocaleDateString();
    }
  }
  
  return strValue;
}

export function ExcelTable({ 
  siteName, 
  folderPath = '', 
  title = "Excel Spreadsheets",
  description = "View and interact with Excel spreadsheets from SharePoint",
  pageSize = 25
}: ExcelTableProps) {
  const [selectedFile, setSelectedFile] = useState<SharePointExcelFile | null>(null);
  const [selectedWorksheet, setSelectedWorksheet] = useState<string>('');
  const [currentPage, setCurrentPage] = useState(1);
  const [searchTerm, setSearchTerm] = useState('');

  // Query for listing Excel files
  const { 
    data: filesData, 
    isLoading: filesLoading, 
    error: filesError,
    refetch: refetchFiles
  } = useQuery({
    queryKey: ['/api/sharepoint-excel/files', siteName, folderPath],
    queryFn: async () => {
      try {
        const token = await authService.getAccessToken();
        const params = new URLSearchParams({
          site: siteName,
          ...(folderPath && { folder: folderPath })
        });
        
        const response = await fetch(`/api/sharepoint-excel/files?${params}`, {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json'
          },
          credentials: 'include'
        });
        
        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to fetch Excel files: ${response.status} - ${errorText}`);
        }
        
        return await response.json();
      } catch (error) {
        console.error('Error fetching Excel files:', error);
        throw error;
      }
    },
    enabled: true,
    retry: false,
    staleTime: 5 * 60 * 1000 // 5 minutes
  });

  // Query for Excel data when file is selected - fetch all worksheets once
  const { 
    data: workbookData, 
    isLoading: dataLoading, 
    error: dataError,
    refetch: refetchData
  } = useQuery({
    queryKey: ['/api/sharepoint-excel/data', siteName, folderPath, selectedFile?.id || selectedFile?.name],
    queryFn: async () => {
      if (!selectedFile) return null;
      
      try {
        const token = await authService.getAccessToken();
        const params = new URLSearchParams({
          site: siteName,
          fileName: selectedFile.name,
          ...(folderPath && { folder: folderPath })
          // No worksheet param - fetch all sheets at once
        });
        
        const response = await fetch(`/api/sharepoint-excel/data?${params}`, {
          headers: {
            'Authorization': `Bearer ${token}`,
            'Accept': 'application/json'
          },
          credentials: 'include'
        });
        
        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Failed to fetch Excel data: ${response.status} - ${errorText}`);
        }
        
        return await response.json();
      } catch (error) {
        console.error('Error fetching Excel data:', error);
        throw error;
      }
    },
    enabled: !!selectedFile,
    retry: false,
    refetchOnWindowFocus: false,
    refetchOnReconnect: false,
    staleTime: 15 * 60 * 1000 // 15 minutes
  });

  const handleFileSelect = (fileName: string) => {
    const file = filesData?.files?.find((f: SharePointExcelFile) => f.name === fileName);
    if (file) {
      setSelectedFile(file);
      setSelectedWorksheet(''); // Reset worksheet selection
      setCurrentPage(1); // Reset pagination
    }
  };

  const handleWorksheetSelect = (worksheetName: string) => {
    setSelectedWorksheet(worksheetName);
    setCurrentPage(1); // Reset pagination when changing worksheets
  };

  const handleRefresh = () => {
    refetchFiles();
    if (selectedFile) {
      refetchData();
    }
  };

  // Auto-select worksheet when there's only one available
  useEffect(() => {
    if (workbookData?.data?.sheets && workbookData.data.sheets.length === 1 && !selectedWorksheet) {
      const singleWorksheet = workbookData.data.sheets[0];
      setSelectedWorksheet(singleWorksheet.name);
    }
  }, [workbookData, selectedWorksheet]);


  const handleDownloadExcel = () => {
    if (selectedFile) {
      // Open directly in SharePoint - secure and handles authentication properly
      window.open(selectedFile.webUrl, '_blank');
    }
  };

  // Get current worksheet data
  const currentWorksheet = workbookData?.data?.sheets?.find(
    (sheet: WorksheetData) => selectedWorksheet ? sheet.name === selectedWorksheet : true
  ) || workbookData?.data?.sheets?.[0];

  // Filter rows based on search term
  const filteredRows = useMemo(() => {
    if (!currentWorksheet?.rows || !searchTerm.trim()) {
      return currentWorksheet?.rows || [];
    }
    
    const term = searchTerm.toLowerCase();
    return currentWorksheet.rows.filter((row: (string | number | null)[]) => {
      return row.some((cell, index) => {
        if (cell === null || cell === undefined) return false;
        const header = currentWorksheet.headers[index] || '';
        const formattedValue = formatCellValue(cell, header);
        return formattedValue.toLowerCase().includes(term);
      });
    });
  }, [currentWorksheet, searchTerm]);

  // Pagination logic
  const totalRows = filteredRows.length;
  const totalPages = Math.ceil(totalRows / pageSize);
  const startIndex = (currentPage - 1) * pageSize;
  const endIndex = startIndex + pageSize;
  const paginatedRows = filteredRows.slice(startIndex, endIndex);

  // Reset to page 1 when search changes
  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm]);

  return (
    <div className="space-y-6" data-testid="excel-table-container">
      <Card>
        <CardHeader>
          <div className="flex items-center justify-between">
            <div>
              <CardTitle className="flex items-center gap-2">
                <FileSpreadsheet className="h-5 w-5 text-green-600" />
                {title}
              </CardTitle>
              <CardDescription className="mt-2">{description}</CardDescription>
            </div>
            <div className="flex gap-2">
              <Button
                variant="outline"
                size="sm"
                onClick={handleRefresh}
                disabled={filesLoading || dataLoading}
                data-testid="button-refresh"
              >
                <RefreshCw className={`h-4 w-4 mr-2 ${(filesLoading || dataLoading) ? 'animate-spin' : ''}`} />
                Reload Data
              </Button>
            </div>
          </div>
        </CardHeader>
        <CardContent className="space-y-4">
          {/* File Selection */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-2">
              <label className="text-sm font-medium">Select Excel File</label>
              {filesLoading ? (
                <Skeleton className="h-10 w-full" />
              ) : filesError ? (
                <Alert variant="destructive">
                  <AlertCircle className="h-4 w-4" />
                  <AlertDescription>
                    Failed to load Excel files: {filesError instanceof Error ? filesError.message : 'Unknown error'}
                  </AlertDescription>
                </Alert>
              ) : (
                <Select onValueChange={handleFileSelect} value={selectedFile?.name || ''}>
                  <SelectTrigger data-testid="select-excel-file">
                    <SelectValue placeholder="Choose an Excel file..." />
                  </SelectTrigger>
                  <SelectContent>
                    {filesData?.files?.map((file: SharePointExcelFile) => (
                      <SelectItem key={file.id} value={file.name}>
                        <div className="flex items-center gap-2">
                          <Badge variant="secondary" className="text-xs">
                            {file.type.toUpperCase()}
                          </Badge>
                          <span className="truncate">{file.name}</span>
                        </div>
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              )}
            </div>

            {/* Worksheet Selection */}
            {selectedFile && workbookData?.data && (
              <div className="space-y-2">
                <label className="text-sm font-medium">Select Worksheet</label>
                <Select onValueChange={handleWorksheetSelect} value={selectedWorksheet}>
                  <SelectTrigger data-testid="select-worksheet">
                    <SelectValue placeholder="Choose a worksheet..." />
                  </SelectTrigger>
                  <SelectContent>
                    {workbookData.data.sheets.map((sheet: WorksheetData) => (
                      <SelectItem key={sheet.name} value={sheet.name}>
                        {sheet.name} ({sheet.rows.length} items)
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            )}
          </div>

          {/* File Info and Actions */}
          {selectedFile && (
            <div className="flex items-center justify-between p-3 bg-gray-50 rounded-lg">
              <div className="flex items-center gap-4">
                <div>
                  <div className="font-medium">{selectedFile.name}</div>
                  <div className="text-sm text-gray-600">
                    {Math.round(selectedFile.size / 1024)} KB • Modified: {new Date(selectedFile.lastModified).toLocaleDateString()}
                  </div>
                </div>
                {workbookData?.data && (
                  <Badge variant="outline">
                    {workbookData.data.sheets.length} worksheet{workbookData.data.sheets.length !== 1 ? 's' : ''}
                  </Badge>
                )}
              </div>
              <Button
                variant="outline"
                size="sm"
                onClick={handleDownloadExcel}
                data-testid="button-download-excel"
              >
                <Download className="h-4 w-4 mr-2" />
                Open in SharePoint
              </Button>
            </div>
          )}
        </CardContent>
      </Card>

      {/* Data Display */}
      {selectedFile && (
        <Card>
          <CardHeader>
            <div className="flex items-center justify-between">
              <div>
                <CardTitle>
                  {currentWorksheet?.name || 'Worksheet Data'}
                </CardTitle>
                <CardDescription>
                  {totalRows > 0 && (
                    <>
                      {currentWorksheet?.headers?.length || 0} columns • 
                      {searchTerm ? (
                        <>{totalRows} of {currentWorksheet?.rows?.length || 0} items</>
                      ) : (
                        <>{totalRows} items</>
                      )}
                      {totalPages > 1 && (
                        <> • Page {currentPage} of {totalPages}</>
                      )}
                    </>
                  )}
                </CardDescription>
              </div>
            </div>
          </CardHeader>
          <CardContent>
            {dataLoading ? (
              <div className="flex items-center justify-center py-12">
                <div className="text-center">
                  <RefreshCw className="h-8 w-8 animate-spin mx-auto mb-4 text-blue-500" />
                  <div className="text-lg font-medium">Loading equipment data...</div>
                  <div className="text-sm text-gray-500 mt-1">Please wait while we fetch your spreadsheet</div>
                </div>
              </div>
            ) : dataError ? (
              <Alert variant="destructive">
                <AlertCircle className="h-4 w-4" />
                <AlertDescription>
                  Failed to load Excel data: {dataError instanceof Error ? dataError.message : 'Unknown error'}
                </AlertDescription>
              </Alert>
            ) : currentWorksheet && currentWorksheet.headers.length > 0 ? (
              <div className="space-y-4">
                {/* Search Filter */}
                <div className="flex items-center space-x-2">
                  <Search className="h-4 w-4 text-gray-400" />
                  <Input
                    placeholder="Search equipment data..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="max-w-sm"
                    data-testid="search-input"
                  />
                  {searchTerm && (
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => setSearchTerm('')}
                      data-testid="clear-search"
                    >
                      Clear
                    </Button>
                  )}
                </div>
                <div className="overflow-x-auto">
                  <Table data-testid="excel-data-table">
                    <TableHeader>
                      <TableRow>
                        {currentWorksheet.headers.map((header: string, index: number) => (
                          <TableHead key={index} className="font-semibold">
                            {header || `Column ${index + 1}`}
                          </TableHead>
                        ))}
                      </TableRow>
                    </TableHeader>
                    <TableBody>
                      {paginatedRows.map((row: (string | number | null)[], rowIndex: number) => (
                        <TableRow key={startIndex + rowIndex} data-testid={`row-excel-${startIndex + rowIndex}`}>
                          {currentWorksheet.headers.map((header: string, colIndex: number) => {
                            const cellValue = row[colIndex];
                            const isUrl = cellValue && typeof cellValue === 'string' && 
                              (cellValue.startsWith('http://') || cellValue.startsWith('https://'));
                            const isImageColumn = header && (
                              header.toLowerCase().includes('qr code') || 
                              header.toLowerCase().includes('photo') || 
                              header.toLowerCase().includes('image') ||
                              header.toLowerCase().includes('picture') ||
                              header.toLowerCase().includes('url') ||
                              header.toLowerCase().includes('link') ||
                              header.toLowerCase().includes('src') ||
                              header.toLowerCase().includes('href')
                            );
                            
                            // Check if this might be a checkbox column
                            const headerLower = header.toLowerCase();
                            const isCheckboxColumn = header && (
                              headerLower.includes('tagged') ||
                              headerLower.includes('checked') ||
                              headerLower.includes('complete') ||
                              headerLower.includes('done') ||
                              headerLower.includes('active') ||
                              headerLower.includes('enabled') ||
                              headerLower.includes('valid')
                            );
                            
                            const checkboxData = formatCheckboxValue(cellValue);
                            
                            return (
                              <TableCell key={colIndex}>
                                {cellValue !== null && cellValue !== undefined ? (
                                  (isUrl && isImageColumn) ? (
                                    <ImageCell 
                                      url={cellValue} 
                                      altText={cellValue}
                                      columnName={header}
                                    />
                                  ) : (isCheckboxColumn && checkboxData.isCheckbox) ? (
                                    <div className="flex items-center justify-center" data-testid={`checkbox-${checkboxData.checked ? 'checked' : 'unchecked'}`}>
                                      {checkboxData.checked ? (
                                        <Check className="h-4 w-4 text-green-600" />
                                      ) : (
                                        <X className="h-4 w-4 text-red-500" />
                                      )}
                                    </div>
                                  ) : (
                                    formatCellValue(cellValue, header)
                                  )
                                ) : (
                                  ''
                                )}
                              </TableCell>
                            );
                          })}
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </div>

                {/* Pagination */}
                {totalPages > 1 && (
                  <div className="flex items-center justify-between">
                    <div className="text-sm text-gray-600">
                      Showing {startIndex + 1} to {Math.min(endIndex, totalRows)} of {totalRows} items
                    </div>
                    <Pagination>
                      <PaginationContent>
                        <PaginationItem>
                          <PaginationPrevious
                            onClick={() => setCurrentPage(Math.max(1, currentPage - 1))}
                            className={currentPage === 1 ? 'pointer-events-none opacity-50' : 'cursor-pointer'}
                            data-testid="button-prev-page"
                          />
                        </PaginationItem>
                        <PaginationItem>
                          <span className="px-3 py-2 text-sm">
                            Page {currentPage} of {totalPages}
                          </span>
                        </PaginationItem>
                        <PaginationItem>
                          <PaginationNext
                            onClick={() => setCurrentPage(Math.min(totalPages, currentPage + 1))}
                            className={currentPage === totalPages ? 'pointer-events-none opacity-50' : 'cursor-pointer'}
                            data-testid="button-next-page"
                          />
                        </PaginationItem>
                      </PaginationContent>
                    </Pagination>
                  </div>
                )}
              </div>
            ) : (
              <div className="text-center py-8">
                <div className="text-gray-500 mb-2">No data available</div>
                <div className="text-sm text-gray-400">
                  {currentWorksheet ? 'This worksheet appears to be empty' : 'Select a file and worksheet to view data'}
                </div>
              </div>
            )}
          </CardContent>
        </Card>
      )}
    </div>
  );
}