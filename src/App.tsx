/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { Settings, AlertCircle, Loader2, Table, LogIn, Share2, Check, Maximize, Minimize } from 'lucide-react';

interface DataRow {
  id: string;
  name: string;
  dates: string;
}

interface PageData {
  sheetName: string;
  items: DataRow[];
}

const ColumnInput = ({ label, value, onChange, placeholder, required = false, availableHeaders = [] }: any) => {
  const [mode, setMode] = useState<'select' | 'input'>(
    availableHeaders.length > 0 && availableHeaders.includes(value) ? 'select' : (availableHeaders.length > 0 && !value ? 'select' : 'input')
  );

  useEffect(() => {
    if (availableHeaders.length > 0 && availableHeaders.includes(value)) {
      setMode('select');
    } else if (availableHeaders.length === 0) {
      setMode('input');
    }
  }, [availableHeaders, value]);

  return (
    <div>
      <label className="block text-sm font-semibold text-gray-700 mb-1">
        {label}
      </label>
      {mode === 'select' && availableHeaders.length > 0 ? (
        <select
          value={value || ''}
          onChange={(e) => {
            if (e.target.value === '__custom__') {
              setMode('input');
              onChange('');
            } else {
              onChange(e.target.value);
            }
          }}
          className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all bg-white"
          required={required}
        >
          <option value="" disabled>Select a column</option>
          {availableHeaders.map((h: string, i: number) => (
            <option key={i} value={h}>{h}</option>
          ))}
          <option value="__custom__">Type custom column letter...</option>
        </select>
      ) : (
        <div className="relative">
          <input
            type="text"
            value={value}
            onChange={(e) => onChange(e.target.value)}
            placeholder={placeholder}
            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all pr-20"
            required={required}
          />
          {availableHeaders.length > 0 && (
            <button
              type="button"
              onClick={() => {
                setMode('select');
                onChange(availableHeaders[0] || '');
              }}
              className="absolute right-3 top-1/2 -translate-y-1/2 text-xs text-[#000066] hover:underline font-medium"
            >
              Use List
            </button>
          )}
        </div>
      )}
      <p className="text-xs text-gray-500 mt-1">Select header or type column letter</p>
    </div>
  );
};

export default function App() {
  const params = new URLSearchParams(window.location.search);
  
  const [isConfiguring, setIsConfiguring] = useState(params.get('auto') !== 'true');
  const [sheetUrl, setSheetUrl] = useState(params.get('url') || '');
  const [nameColInput, setNameColInput] = useState(params.get('name') || 'Name');
  const [datesColInput, setDatesColInput] = useState(params.get('dates') || 'Dates');
  const [startDateColInput, setStartDateColInput] = useState(params.get('start') || 'Start Date');
  const [endDateColInput, setEndDateColInput] = useState(params.get('end') || 'End Date');
  const [headerRowInput, setHeaderRowInput] = useState(params.get('header') || '2');
  const [itemsPerPageInput, setItemsPerPageInput] = useState(params.get('items') || '8');
  const [pageDurationInput, setPageDurationInput] = useState(params.get('duration') || '10');
  const [refreshIntervalInput, setRefreshIntervalInput] = useState(params.get('refresh') || '5');
  const [currentPage, setCurrentPage] = useState(0);
  const [pages, setPages] = useState<PageData[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [displayTitle, setDisplayTitle] = useState(params.get('title') || 'Display Examples');
  
  const [availableSheets, setAvailableSheets] = useState<{sheetId: number, title: string}[]>([]);
  const [selectedSheetTitles, setSelectedSheetTitles] = useState<string[]>(
    params.get('sheets') ? params.get('sheets')!.split(',').map(s => s.trim()).filter(Boolean) : []
  );
  const [availableHeaders, setAvailableHeaders] = useState<string[]>([]);
  const [previewData, setPreviewData] = useState<any[][]>([]);
  const [isFetchingMetadata, setIsFetchingMetadata] = useState(false);
  const [copied, setCopied] = useState(false);
  const [isFullscreen, setIsFullscreen] = useState(false);
  
  const [authStatus, setAuthStatus] = useState({ isAuthenticated: false, isConfigured: false, checking: true });

  useEffect(() => {
    const handleFullscreenChange = () => {
      setIsFullscreen(!!document.fullscreenElement);
    };
    document.addEventListener('fullscreenchange', handleFullscreenChange);
    return () => document.removeEventListener('fullscreenchange', handleFullscreenChange);
  }, []);

  useEffect(() => {
    const params = new URLSearchParams();
    if (sheetUrl) params.set('url', sheetUrl);
    if (nameColInput !== 'Name') params.set('name', nameColInput);
    if (datesColInput !== 'Dates') params.set('dates', datesColInput);
    if (startDateColInput !== 'Start Date') params.set('start', startDateColInput);
    if (endDateColInput !== 'End Date') params.set('end', endDateColInput);
    if (headerRowInput !== '2') params.set('header', headerRowInput);
    if (itemsPerPageInput !== '8') params.set('items', itemsPerPageInput);
    if (pageDurationInput !== '10') params.set('duration', pageDurationInput);
    if (refreshIntervalInput !== '5') params.set('refresh', refreshIntervalInput);
    if (displayTitle !== 'Display Examples') params.set('title', displayTitle);
    if (selectedSheetTitles.length > 0) params.set('sheets', selectedSheetTitles.map(s => s.trim()).join(','));
    if (!isConfiguring) params.set('auto', 'true');
    
    const newUrl = `${window.location.pathname}?${params.toString()}`;
    window.history.replaceState({}, '', newUrl);
  }, [
    sheetUrl, nameColInput, datesColInput, startDateColInput, endDateColInput,
    headerRowInput, itemsPerPageInput, pageDurationInput, refreshIntervalInput,
    displayTitle, selectedSheetTitles, isConfiguring
  ]);

  useEffect(() => {
    if (!authStatus.checking && !authStatus.isAuthenticated && !isConfiguring) {
      setIsConfiguring(true);
    }
  }, [authStatus.checking, authStatus.isAuthenticated, isConfiguring]);

  useEffect(() => {
    if (authStatus.isAuthenticated && !isConfiguring && pages.length === 0 && !loading && !error) {
      fetchData(false);
    }
  }, [authStatus.isAuthenticated, isConfiguring, pages.length, loading, error]);

  useEffect(() => {
    if (!previewData || previewData.length === 0) return;
    const headerRowIdx = Math.max(0, parseInt(headerRowInput) - 1 || 0);
    const headers = previewData[headerRowIdx] ? previewData[headerRowIdx].map((h: string) => h.trim()) : [];
    const validHeaders = headers.filter(Boolean);
    setAvailableHeaders(validHeaders);
    
    if (validHeaders.includes('Name')) setNameColInput('Name');
    if (validHeaders.includes('Dates')) setDatesColInput('Dates');
    if (validHeaders.includes('Start Date')) setStartDateColInput('Start Date');
    if (validHeaders.includes('End Date')) setEndDateColInput('End Date');
  }, [previewData, headerRowInput]);

  useEffect(() => {
    if (isConfiguring || pages.length === 0) return;
    
    const totalPages = pages.length;
    
    if (totalPages <= 1) return;

    const durationMs = (parseInt(pageDurationInput) || 20) * 1000;
    
    const interval = setInterval(() => {
      setCurrentPage((prev) => (prev + 1) % totalPages);
    }, durationMs);

    return () => clearInterval(interval);
  }, [isConfiguring, pages.length, pageDurationInput]);

  useEffect(() => {
    if (isConfiguring) return;

    const intervalMinutes = parseFloat(refreshIntervalInput);
    if (isNaN(intervalMinutes) || intervalMinutes <= 0) return;

    const intervalMs = intervalMinutes * 60 * 1000;
    const interval = setInterval(() => {
      fetchData(true);
    }, intervalMs);

    return () => clearInterval(interval);
  }, [isConfiguring, refreshIntervalInput, sheetUrl, nameColInput, datesColInput, startDateColInput, endDateColInput, headerRowInput]);

  useEffect(() => {
    checkAuthStatus();
    const handleMessage = (event: MessageEvent) => {
      const origin = event.origin;
      if (!origin.endsWith('.run.app') && !origin.includes('localhost')) return;
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        checkAuthStatus();
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const checkAuthStatus = async () => {
    try {
      const res = await fetch('/api/auth/status');
      const data = await res.json();
      setAuthStatus({ ...data, checking: false });
    } catch (e) {
      setAuthStatus(prev => ({ ...prev, checking: false }));
    }
  };

  const handleConnectGoogle = async () => {
    try {
      const response = await fetch('/api/auth/url');
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Failed to get auth URL');

      const authWindow = window.open(data.url, 'oauth_popup', 'width=600,height=700');
      if (!authWindow) {
        alert('Please allow popups for this site to connect your account.');
      }
    } catch (error: any) {
      setError(error.message || 'Failed to initiate Google Sign-In.');
    }
  };

  const extractSheetId = (input: string) => {
    const match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : input;
  };

  const extractGid = (input: string) => {
    const match = input.match(/[#&]gid=([0-9]+)/);
    return match ? match[1] : '';
  };

  const fetchMetadataAndData = async () => {
    setError('');
    setIsFetchingMetadata(true);
    try {
      const sheetId = extractSheetId(sheetUrl);
      if (!sheetId) throw new Error('Please enter a valid Google Sheet URL or ID.');

      const metaRes = await fetch(`/api/sheet/metadata?sheetId=${sheetId}`);
      const metaResult = await metaRes.json();
      if (!metaRes.ok) {
        if (metaRes.status === 401) {
          setAuthStatus(prev => ({ ...prev, isAuthenticated: false }));
          throw new Error('Authentication expired. Please sign in again.');
        }
        throw new Error(metaResult.error || 'Failed to fetch metadata.');
      }

      setAvailableSheets(metaResult.sheets);

      const gid = extractGid(sheetUrl);
      let targetSheet = metaResult.sheets[0]?.title;
      if (gid) {
        const found = metaResult.sheets.find((s: any) => s.sheetId === Number(gid));
        if (found) targetSheet = found.title;
      }
      
      const validExistingSheets = selectedSheetTitles.filter(title => 
        metaResult.sheets.some((s: any) => s.title === title)
      );

      if (validExistingSheets.length > 0) {
        if (validExistingSheets.length !== selectedSheetTitles.length || !validExistingSheets.every((v, i) => v === selectedSheetTitles[i])) {
          setSelectedSheetTitles(validExistingSheets);
        }
        await fetchHeadersForSheet(sheetId, validExistingSheets[0]);
      } else {
        setSelectedSheetTitles(targetSheet ? [targetSheet] : []);
        if (targetSheet) {
          await fetchHeadersForSheet(sheetId, targetSheet);
        }
      }
    } catch (err: any) {
      setError(err.message || 'Failed to fetch sheets.');
    } finally {
      setIsFetchingMetadata(false);
    }
  };

  useEffect(() => {
    if (isConfiguring && sheetUrl && availableSheets.length === 0 && !isFetchingMetadata && authStatus.isAuthenticated) {
      fetchMetadataAndData();
    }
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isConfiguring, authStatus.isAuthenticated]);

  const fetchHeadersForSheet = async (sheetId: string, sheetName: string) => {
    try {
      const response = await fetch(`/api/sheet?sheetId=${sheetId}&sheetName=${encodeURIComponent(sheetName)}`);
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Failed to fetch sheet data.');
      
      const values: any[][] = result.values;
      if (!values || values.length === 0) return;

      setPreviewData(values.slice(0, 10)); // Store first 10 rows for preview/header detection
    } catch (err) {
      console.error("Failed to fetch headers", err);
    }
  };

  const fetchData = async (isBackground = false) => {
    if (!isBackground) {
      setError('');
      setLoading(true);
    }

    try {
      const sheetId = extractSheetId(sheetUrl);
      const gid = extractGid(sheetUrl);
      
      if (!sheetId) throw new Error('Please enter a valid Google Sheet URL or ID.');

      const sheetsToFetch = selectedSheetTitles.length > 0 ? selectedSheetTitles : [null];
      const allPages: PageData[] = [];
      let firstResolvedName = '';
      const itemsPerPage = parseInt(itemsPerPageInput) || 8;

      for (const sheetName of sheetsToFetch) {
        let query = `?sheetId=${sheetId}`;
        if (sheetName) {
          query += `&sheetName=${encodeURIComponent(sheetName)}`;
        } else if (gid) {
          query += `&gid=${gid}`;
        }

        const response = await fetch(`/api/sheet${query}`);
        const result = await response.json();

        if (!response.ok) {
          if (response.status === 401) {
            setAuthStatus(prev => ({ ...prev, isAuthenticated: false }));
            if (!isBackground) setIsConfiguring(true);
            throw new Error('Authentication expired. Please sign in again.');
          }
          throw new Error(result.error || 'Failed to fetch data.');
        }

        if (!firstResolvedName) firstResolvedName = result.resolvedSheetName;

        const values: any[][] = result.values || [];
        const parsedData: DataRow[] = [];

        if (values.length > 0) {
          const headerRowIdx = Math.max(0, parseInt(headerRowInput) - 1 || 0);
          const startRow = headerRowIdx + 2;

          const headers = values[headerRowIdx] ? values[headerRowIdx].map((h: string) => h.toLowerCase().trim()) : [];

          const resolveColIndex = (input: string) => {
            const trimmed = input.trim();
            if (!trimmed) return -1;
            
            // Try exact header match first (case-insensitive)
            const idx = headers.findIndex((h: string) => h === trimmed.toLowerCase());
            if (idx !== -1) return idx;

            // If not found as a header, check if it's a valid column letter
            if (/^[A-Z]+$/i.test(trimmed)) {
              let colIdx = 0;
              const upper = trimmed.toUpperCase();
              for (let i = 0; i < upper.length; i++) {
                colIdx = colIdx * 26 + (upper.charCodeAt(i) - 64);
              }
              return colIdx - 1;
            }

            return -1;
          };

          const nameIdx = resolveColIndex(nameColInput);
          const datesIdx = resolveColIndex(datesColInput);
          const startDateIdx = resolveColIndex(startDateColInput);
          const endDateIdx = resolveColIndex(endDateColInput);

          if (nameIdx === -1 && datesIdx === -1) {
            console.warn(`Could not find columns matching "${nameColInput}" or "${datesColInput}" in sheet ${result.resolvedSheetName}.`);
          } else {
            for (let i = startRow - 1; i < values.length; i++) {
              const row = values[i];
              if (!row) continue;

              // Filter out rows with start date in the future
              if (startDateIdx !== -1 && startDateIdx < row.length) {
                const startDateStr = row[startDateIdx];
                if (startDateStr) {
                  const startDate = new Date(startDateStr);
                  if (!isNaN(startDate.getTime())) {
                    if (startDate.getTime() > Date.now()) {
                      continue; // Skip this row as it starts in the future
                    }
                  }
                }
              }

              if (endDateIdx !== -1 && endDateIdx < row.length) {
                const endDateStr = row[endDateIdx];
                if (endDateStr) {
                  const endDate = new Date(endDateStr);
                  if (!isNaN(endDate.getTime())) {
                    // If no specific time was provided (exactly midnight), push to the end of the day
                    if (endDate.getHours() === 0 && endDate.getMinutes() === 0 && endDate.getSeconds() === 0) {
                      endDate.setHours(23, 59, 59, 999);
                    }
                    if (endDate.getTime() < Date.now()) {
                      continue; // Skip this row as it has passed the end date
                    }
                  }
                }
              }

              const name = nameIdx !== -1 && nameIdx < row.length ? row[nameIdx] : '';
              const dates = datesIdx !== -1 && datesIdx < row.length ? row[datesIdx] : '';
              if (name || dates) {
                parsedData.push({ id: `${result.resolvedSheetName}-${i}`, name, dates });
              }
            }
          }
        }
        
        if (parsedData.length > 0) {
          const sheetPages = Math.ceil(parsedData.length / itemsPerPage);
          for (let p = 0; p < sheetPages; p++) {
            allPages.push({
              sheetName: result.resolvedSheetName,
              items: parsedData.slice(p * itemsPerPage, (p + 1) * itemsPerPage)
            });
          }
        } else {
          allPages.push({
            sheetName: result.resolvedSheetName,
            items: []
          });
        }
      }

      setPages(allPages);
      setDisplayTitle(firstResolvedName || 'Sheet Data');
      if (!isBackground) {
        setCurrentPage(0);
        setIsConfiguring(false);
      }
    } catch (err: any) {
      if (!isBackground) {
        setError(err.message || 'An unexpected error occurred.');
      } else {
        console.error('Background refresh failed:', err);
      }
    } finally {
      if (!isBackground) {
        setLoading(false);
      }
    }
  };

  const loadData = async (e: React.FormEvent) => {
    e.preventDefault();
    await fetchData(false);
  };

  const handleCopyLink = () => {
    navigator.clipboard.writeText(window.location.href);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const toggleFullscreen = async () => {
    if (!document.fullscreenElement) {
      try {
        await document.documentElement.requestFullscreen();
      } catch (err) {
        console.error("Error attempting to enable fullscreen:", err);
      }
    } else {
      if (document.exitFullscreen) {
        await document.exitFullscreen();
      }
    }
  };

  if (authStatus.checking) {
    return (
      <div className="min-h-screen bg-[#000066] flex items-center justify-center">
        <Loader2 size={48} className="animate-spin text-[#ffcc66]" />
      </div>
    );
  }

  if (isConfiguring) {
    return (
      <div className="min-h-screen bg-[#000066] p-8 flex items-center justify-center font-sans">
        <div className="bg-white rounded-2xl p-8 max-w-md w-full shadow-2xl">
          <div className="flex items-center gap-3 mb-6">
            <div className="bg-[#000066] p-3 rounded-xl text-white">
              <Table size={24} />
            </div>
            <h2 className="text-2xl font-bold text-[#000066]">Connect Sheet</h2>
          </div>

          {!authStatus.isConfigured ? (
            <div className="bg-blue-50 text-blue-800 p-4 rounded-lg text-sm mb-6">
              <h3 className="font-bold mb-2 flex items-center gap-2">
                <AlertCircle size={16} /> Setup Required
              </h3>
              <p className="mb-2">To connect to private sheets, you need to configure Google OAuth.</p>
              <ol className="list-decimal list-inside space-y-1 ml-1">
                <li>Create a Google Cloud Project</li>
                <li>Enable the Google Sheets API</li>
                <li>Configure OAuth Consent Screen</li>
                <li>Create OAuth Client ID (Web Application)</li>
                <li>Add the Redirect URI shown below</li>
                <li>Set <strong>GOOGLE_CLIENT_ID</strong> and <strong>GOOGLE_CLIENT_SECRET</strong> in AI Studio settings</li>
              </ol>
            </div>
          ) : !authStatus.isAuthenticated ? (
            <div className="mb-6">
              <p className="text-sm text-gray-600 mb-4">
                Sign in with Google to access your private sheets.
              </p>
              <button
                onClick={handleConnectGoogle}
                className="w-full bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 font-semibold py-3 px-4 rounded-lg transition-colors flex items-center justify-center gap-3"
              >
                <LogIn size={20} />
                Sign in with Google
              </button>
              {error && (
                <div className="mt-4 bg-red-50 text-red-600 p-3 rounded-lg flex items-start gap-2 text-sm">
                  <AlertCircle size={16} className="mt-0.5 shrink-0" />
                  <span>{error}</span>
                </div>
              )}
            </div>
          ) : (
            <form onSubmit={loadData} className="space-y-5">
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1">
                  Google Sheet URL or ID
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={sheetUrl}
                    onChange={(e) => setSheetUrl(e.target.value)}
                    placeholder="https://docs.google.com/spreadsheets/d/..."
                    className="flex-1 px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                    required
                  />
                  <button
                    type="button"
                    onClick={fetchMetadataAndData}
                    disabled={isFetchingMetadata || !sheetUrl}
                    className="bg-gray-100 hover:bg-gray-200 text-gray-700 font-semibold py-2 px-4 rounded-lg transition-colors disabled:opacity-50 flex items-center justify-center min-w-[100px]"
                  >
                    {isFetchingMetadata ? <Loader2 size={20} className="animate-spin" /> : 'Fetch'}
                  </button>
                </div>
              </div>

              {availableSheets.length > 0 && (
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-2">
                    Select Sheets
                  </label>
                  <div className="max-h-40 overflow-y-auto border border-gray-300 rounded-lg p-2 space-y-1 bg-white">
                    {availableSheets.map(s => (
                      <label key={s.sheetId} className="flex items-center gap-3 p-2 hover:bg-gray-50 rounded cursor-pointer transition-colors">
                        <input
                          type="checkbox"
                          checked={selectedSheetTitles.includes(s.title)}
                          onChange={(e) => {
                            const newSelection = e.target.checked
                              ? [...selectedSheetTitles, s.title]
                              : selectedSheetTitles.filter(t => t !== s.title);
                            
                            const oldFirst = selectedSheetTitles[0];
                            const newFirst = newSelection[0];
                            
                            setSelectedSheetTitles(newSelection);
                            
                            if (newFirst && newFirst !== oldFirst) {
                              fetchHeadersForSheet(extractSheetId(sheetUrl), newFirst);
                            } else if (!newFirst) {
                              setPreviewData([]);
                            }
                          }}
                          className="w-4 h-4 text-[#000066] rounded border-gray-300 focus:ring-[#000066]"
                        />
                        <span className="text-sm text-gray-700 font-medium">{s.title}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}

              <div className="grid grid-cols-2 gap-4">
                <ColumnInput
                  label="Name Column"
                  value={nameColInput}
                  onChange={setNameColInput}
                  placeholder="e.g., Name or A"
                  required={true}
                  availableHeaders={availableHeaders}
                />
                <ColumnInput
                  label="Dates Column"
                  value={datesColInput}
                  onChange={setDatesColInput}
                  placeholder="e.g., Dates or B"
                  required={true}
                  availableHeaders={availableHeaders}
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <ColumnInput
                  label="Start Date Column (Optional)"
                  value={startDateColInput}
                  onChange={setStartDateColInput}
                  placeholder="e.g., Start Date or C"
                  availableHeaders={availableHeaders}
                />
                <ColumnInput
                  label="End Date Column (Optional)"
                  value={endDateColInput}
                  onChange={setEndDateColInput}
                  placeholder="e.g., End Date or D"
                  availableHeaders={availableHeaders}
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Header Row
                  </label>
                  <input
                    type="number"
                    min="1"
                    value={headerRowInput}
                    onChange={(e) => setHeaderRowInput(e.target.value)}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                    required
                  />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Items per Page
                  </label>
                  <input
                    type="number"
                    min="1"
                    value={itemsPerPageInput}
                    onChange={(e) => setItemsPerPageInput(e.target.value)}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                    required
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Page Duration (seconds)
                  </label>
                  <input
                    type="number"
                    min="1"
                    value={pageDurationInput}
                    onChange={(e) => setPageDurationInput(e.target.value)}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                    required
                  />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Refresh Data (minutes)
                  </label>
                  <input
                    type="number"
                    min="0.1"
                    step="0.1"
                    value={refreshIntervalInput}
                    onChange={(e) => setRefreshIntervalInput(e.target.value)}
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                    required
                  />
                </div>
              </div>

              {error && (
                <div className="bg-red-50 text-red-600 p-3 rounded-lg flex items-start gap-2 text-sm">
                  <AlertCircle size={16} className="mt-0.5 shrink-0" />
                  <span>{error}</span>
                </div>
              )}

              <div className="flex gap-4 pt-4">
                <button
                  type="button"
                  onClick={handleCopyLink}
                  className="flex-1 bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 font-semibold py-3 px-4 rounded-lg transition-colors flex items-center justify-center gap-2"
                >
                  {copied ? <Check size={20} className="text-green-600" /> : <Share2 size={20} />}
                  {copied ? 'Copied!' : 'Copy Link'}
                </button>
                <button
                  type="submit"
                  disabled={loading}
                  className="flex-[2] bg-[#ffcc66] hover:bg-[#ffb833] text-[#000066] font-bold py-3 px-4 rounded-lg transition-colors flex items-center justify-center gap-2 disabled:opacity-70 disabled:cursor-not-allowed"
                >
                  {loading ? (
                    <>
                      <Loader2 size={20} className="animate-spin" />
                      Loading Data...
                    </>
                  ) : (
                    'Load Grid'
                  )}
                </button>
              </div>
            </form>
          )}
        </div>
      </div>
    );
  }

  if (loading && pages.length === 0) {
    return (
      <div className="min-h-screen bg-[#000066] flex flex-col items-center justify-center font-sans">
        <Loader2 size={48} className="animate-spin text-[#ffcc66] mb-4" />
        <p className="text-white text-lg font-medium">Loading Schedule Data...</p>
      </div>
    );
  }

  if (error && pages.length === 0) {
    return (
      <div className="min-h-screen bg-[#000066] flex flex-col items-center justify-center font-sans p-8 text-center">
        <div className="bg-white p-8 rounded-2xl max-w-md w-full shadow-2xl">
          <AlertCircle size={48} className="text-red-500 mx-auto mb-4" />
          <h2 className="text-2xl font-bold text-gray-800 mb-2">Failed to Load</h2>
          <p className="text-gray-600 mb-6">{error}</p>
          <button
            onClick={() => setIsConfiguring(true)}
            className="w-full bg-[#000066] hover:bg-[#00004d] text-white font-bold py-3 px-4 rounded-lg transition-colors"
          >
            Return to Configuration
          </button>
        </div>
      </div>
    );
  }

  const totalPages = pages.length;
  const currentPageData = pages[currentPage];
  const currentData = currentPageData?.items || [];
  const currentTitle = currentPageData?.sheetName || displayTitle;

  return (
    <div className="min-h-screen bg-[#000066] p-8 md:p-16 font-sans relative flex flex-col">
      <div className="absolute top-6 right-6 md:top-8 md:right-8 flex gap-3 z-10">
        <button
          onClick={toggleFullscreen}
          className="bg-white/10 hover:bg-white/20 text-white p-3 rounded-full transition-colors backdrop-blur-sm"
          title={isFullscreen ? "Exit Fullscreen" : "Enter Fullscreen"}
        >
          {isFullscreen ? <Minimize size={24} /> : <Maximize size={24} />}
        </button>
        <button
          onClick={() => setIsConfiguring(true)}
          className="bg-white/10 hover:bg-white/20 text-white p-3 rounded-full transition-colors backdrop-blur-sm"
          title="Configure Sheet"
        >
          <Settings size={24} />
        </button>
      </div>

      <h1 className="text-5xl md:text-6xl font-bold text-center text-[#ffcc66] mb-16">
        {currentTitle}
      </h1>
      
      <div className="flex-grow">
        <div className="max-w-7xl mx-auto grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-x-12 gap-y-16">
          {currentData.map((row) => (
            <div key={row.id} className="flex flex-col items-center">
            <div className="w-full aspect-[2/1] bg-[#ffcc66] rounded-2xl flex items-center justify-center shadow-lg mb-4 p-4 text-center">
              {row.name && (
                <span className="text-3xl md:text-4xl font-bold text-[#000066] break-words line-clamp-3 whitespace-pre-wrap">
                  {row.name}
                </span>
              )}
            </div>
            {row.dates && (
              <div className="text-center text-white/60 text-xs md:text-sm font-light whitespace-pre-wrap leading-relaxed tracking-wide mt-1">
                {row.dates}
              </div>
            )}
          </div>
        ))}
        </div>
      </div>

      {totalPages > 1 && (
        <div className="mt-12 flex justify-center items-center gap-3">
          {Array.from({ length: totalPages }).map((_, idx) => (
            <div
              key={idx}
              className={`h-2 rounded-full transition-all duration-500 ${
                idx === currentPage ? 'w-8 bg-[#ffcc66]' : 'w-2 bg-white/30'
              }`}
            />
          ))}
        </div>
      )}
    </div>
  );
}


