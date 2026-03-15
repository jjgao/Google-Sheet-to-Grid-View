/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from 'react';
import { Settings, AlertCircle, Loader2, Table, LogIn } from 'lucide-react';

interface DataRow {
  id: string;
  name: string;
  dates: string;
}

interface PageData {
  sheetName: string;
  items: DataRow[];
}

export default function App() {
  const [isConfiguring, setIsConfiguring] = useState(true);
  const [sheetUrl, setSheetUrl] = useState('');
  const [nameColInput, setNameColInput] = useState('Name');
  const [datesColInput, setDatesColInput] = useState('Dates');
  const [startDateColInput, setStartDateColInput] = useState('Start Date');
  const [endDateColInput, setEndDateColInput] = useState('End Date');
  const [headerRowInput, setHeaderRowInput] = useState('2');
  const [itemsPerPageInput, setItemsPerPageInput] = useState('12');
  const [pageDurationInput, setPageDurationInput] = useState('20');
  const [refreshIntervalInput, setRefreshIntervalInput] = useState('5');
  const [currentPage, setCurrentPage] = useState(0);
  const [pages, setPages] = useState<PageData[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [displayTitle, setDisplayTitle] = useState('Display Examples');
  
  const [availableSheets, setAvailableSheets] = useState<{sheetId: number, title: string}[]>([]);
  const [selectedSheetTitles, setSelectedSheetTitles] = useState<string[]>([]);
  const [availableHeaders, setAvailableHeaders] = useState<string[]>([]);
  const [isFetchingMetadata, setIsFetchingMetadata] = useState(false);
  
  const [authStatus, setAuthStatus] = useState({ isAuthenticated: false, isConfigured: false, checking: true });

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
      setSelectedSheetTitles(targetSheet ? [targetSheet] : []);

      if (targetSheet) {
        await fetchHeadersForSheet(sheetId, targetSheet);
      }
    } catch (err: any) {
      setError(err.message || 'Failed to fetch sheets.');
    } finally {
      setIsFetchingMetadata(false);
    }
  };

  const fetchHeadersForSheet = async (sheetId: string, sheetName: string) => {
    try {
      const response = await fetch(`/api/sheet?sheetId=${sheetId}&sheetName=${encodeURIComponent(sheetName)}`);
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || 'Failed to fetch sheet data.');
      
      const values: any[][] = result.values;
      if (!values || values.length === 0) return;

      const headerRowIdx = Math.max(0, parseInt(headerRowInput) - 1 || 0);
      const headers = values[headerRowIdx] ? values[headerRowIdx].map((h: string) => h.trim()) : [];
      const validHeaders = headers.filter(Boolean);
      setAvailableHeaders(validHeaders);
      
      if (validHeaders.includes('Name')) setNameColInput('Name');
      if (validHeaders.includes('Dates')) setDatesColInput('Dates');
      if (validHeaders.includes('Start Date')) setStartDateColInput('Start Date');
      if (validHeaders.includes('End Date')) setEndDateColInput('End Date');
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
      const itemsPerPage = parseInt(itemsPerPageInput) || 12;

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

        const values: any[][] = result.values;
        if (!values || values.length === 0) continue;

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
          continue;
        }

        const parsedData: DataRow[] = [];
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
        
        if (parsedData.length > 0) {
          const sheetPages = Math.ceil(parsedData.length / itemsPerPage);
          for (let p = 0; p < sheetPages; p++) {
            allPages.push({
              sheetName: result.resolvedSheetName,
              items: parsedData.slice(p * itemsPerPage, (p + 1) * itemsPerPage)
            });
          }
        }
      }

      if (allPages.length === 0) {
        throw new Error('No valid data found in the selected sheets.');
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
                            setSelectedSheetTitles(newSelection);
                            if (newSelection.length === 1 && e.target.checked) {
                              fetchHeadersForSheet(extractSheetId(sheetUrl), s.title);
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

              <datalist id="sheet-headers">
                {availableHeaders.map((h, i) => (
                  <option key={i} value={h} />
                ))}
              </datalist>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Name Column
                  </label>
                  <input
                    type="text"
                    list="sheet-headers"
                    value={nameColInput}
                    onChange={(e) => setNameColInput(e.target.value)}
                    placeholder="e.g., Name or A"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                    required
                  />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Dates Column
                  </label>
                  <input
                    type="text"
                    list="sheet-headers"
                    value={datesColInput}
                    onChange={(e) => setDatesColInput(e.target.value)}
                    placeholder="e.g., Dates or B"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                    required
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    Start Date Column (Optional)
                  </label>
                  <input
                    type="text"
                    list="sheet-headers"
                    value={startDateColInput}
                    onChange={(e) => setStartDateColInput(e.target.value)}
                    placeholder="e.g., Start Date or C"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                  />
                </div>
                <div>
                  <label className="block text-sm font-semibold text-gray-700 mb-1">
                    End Date Column (Optional)
                  </label>
                  <input
                    type="text"
                    list="sheet-headers"
                    value={endDateColInput}
                    onChange={(e) => setEndDateColInput(e.target.value)}
                    placeholder="e.g., End Date or D"
                    className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-[#000066] focus:border-[#000066] outline-none transition-all"
                  />
                </div>
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

              <button
                type="submit"
                disabled={loading}
                className="w-full bg-[#ffcc66] hover:bg-[#ffb833] text-[#000066] font-bold py-3 px-4 rounded-lg transition-colors flex items-center justify-center gap-2 disabled:opacity-70 disabled:cursor-not-allowed"
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
            </form>
          )}
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
      <button
        onClick={() => setIsConfiguring(true)}
        className="absolute top-6 right-6 md:top-8 md:right-8 bg-white/10 hover:bg-white/20 text-white p-3 rounded-full transition-colors backdrop-blur-sm z-10"
        title="Configure Sheet"
      >
        <Settings size={24} />
      </button>

      <h1 className="text-5xl md:text-6xl font-bold text-center text-[#ffcc66] mb-16">
        {currentTitle}
      </h1>
      
      <div className="flex-grow">
        <div className="max-w-7xl mx-auto grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-x-12 gap-y-16">
          {currentData.map((row) => (
            <div key={row.id} className="flex flex-col items-center">
            <div className="w-full aspect-[2/1] bg-[#ffcc66] rounded-2xl flex items-center justify-center shadow-lg mb-4 p-4 text-center">
              {row.name && (
                <span className="text-3xl md:text-4xl font-bold text-[#000066] break-words line-clamp-2">
                  {row.name}
                </span>
              )}
            </div>
            {row.dates && (
              <div className="text-center text-white/60 text-xs md:text-sm font-light whitespace-pre-line leading-relaxed tracking-wide mt-1">
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


