import { useState, useEffect, useMemo } from 'react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';
import { LayoutDashboard, LogIn, LogOut, AlertCircle, Loader2, FileSpreadsheet, Filter } from 'lucide-react';
import Papa from 'papaparse';

// Utility to find column index by name (case-insensitive, partial match)
const findColumnIndex = (headers: string[], possibleNames: string[]) => {
  return headers.findIndex(h => 
    possibleNames.some(name => h.toLowerCase().includes(name.toLowerCase()))
  );
};

// Utility to split names by comma or semicolon
const splitNames = (cellValue: string | undefined | null): string[] => {
  if (!cellValue) return [];
  const names = cellValue.toString().split(/[,;]+/).map(n => n.trim()).filter(Boolean);
  return names.length > 0 ? names : [];
};

const extractUnit = (name: string) => {
  const parts = name.split('-');
  if (parts.length > 1) {
    return parts[parts.length - 1].trim();
  }
  return 'Khác';
};

const CustomYAxisTick = (props: any) => {
  const { x, y, payload } = props;
  const text = payload.value;
  const truncated = text.length > 22 ? `${text.substring(0, 22)}...` : text;
  return (
    <g transform={`translate(${x},${y})`}>
      <text x={0} y={0} dy={4} textAnchor="end" fill="#475569" fontSize={12}>
        <title>{text}</title>
        {truncated}
      </text>
    </g>
  );
};

const COLORS = ['#0088FE', '#00C49F', '#FFBB28', '#FF8042', '#8884d8', '#82ca9d', '#ffc658', '#8dd1e1'];

export default function App() {
  const [isLoading, setIsLoading] = useState(true);
  const [data, setData] = useState<{ dataTvkt: any[][], banVeTvkt: any[][] } | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetchData();
  }, []);

  const fetchData = async () => {
    setIsLoading(true);
    setError(null);
    try {
      const spreadsheetId = "1W6r0LnPuQafblFW_7lQ0yLDxjiMuCIrCWkzM3Sg6RkA";
      
      const fetchSheet = async (sheetName: string) => {
        const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
        const response = await fetch(url);
        if (!response.ok) {
          throw new Error(`Failed to fetch sheet ${sheetName}: ${response.statusText}`);
        }
        const csvText = await response.text();
        const parsed = Papa.parse(csvText, { header: false });
        return parsed.data;
      };

      // Try local API first, fallback to direct fetch if it fails or if on static host
      let json;
      try {
        const res = await fetch('/api/sheets/data');
        if (res.ok) {
          json = await res.json();
        } else {
          throw new Error("API not available");
        }
      } catch (e) {
        console.log("Falling back to direct Google Sheets fetch...");
        const dataTvkt = await fetchSheet("Data_TVKT");
        const banVeTvkt = await fetchSheet("Ban_ve_TVKT");
        json = { dataTvkt, banVeTvkt };
      }

      setData(json);
    } catch (err: any) {
      setError(err.message);
    } finally {
      setIsLoading(false);
    }
  };

  if (isLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-slate-50">
        <Loader2 className="w-8 h-8 animate-spin text-indigo-600" />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-indigo-600 text-white rounded-xl flex items-center justify-center">
              <LayoutDashboard className="w-5 h-5" />
            </div>
            <h1 className="text-xl font-bold text-slate-900">dashboard System</h1>
          </div>
          <button 
            onClick={fetchData}
            disabled={isLoading}
            className="flex items-center gap-2 px-4 py-2 bg-indigo-50 text-indigo-700 rounded-lg hover:bg-indigo-100 transition-colors disabled:opacity-50 font-medium"
          >
            <Loader2 className={`w-4 h-4 ${isLoading ? 'animate-spin' : ''}`} />
            {isLoading ? 'Đang cập nhật...' : 'Cập nhật dữ liệu'}
          </button>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {error && (
          <div className="mb-8 p-4 bg-red-50 border border-red-200 text-red-700 rounded-xl flex items-start gap-3">
            <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
            <div>
              <h3 className="font-semibold">Error loading data</h3>
              <p className="text-sm mt-1">{error}</p>
            </div>
          </div>
        )}

        {data && <DashboardContent data={data} onRefresh={fetchData} refreshing={isLoading} />}
      </main>
    </div>
  );
}

function DashboardContent({ data, onRefresh, refreshing }: { 
  data: { dataTvkt: any[][], banVeTvkt: any[][] },
  onRefresh: () => void,
  refreshing: boolean
}) {
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [quickDate, setQuickDate] = useState('all_time');
  const [selectedUnit, setSelectedUnit] = useState('All');
  const [topNTuVan, setTopNTuVan] = useState(10);
  const [topNBanVe, setTopNBanVe] = useState(10);

  const handleQuickDateChange = (value: string) => {
    setQuickDate(value);
    const now = new Date();
    let start = '';
    let end = '';

    const formatDate = (d: Date) => {
      const year = d.getFullYear();
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    };

    switch (value) {
      case 'all_time':
        start = '';
        end = '';
        break;
      case 'this_month':
        start = formatDate(new Date(now.getFullYear(), now.getMonth(), 1));
        end = formatDate(new Date(now.getFullYear(), now.getMonth() + 1, 0));
        break;
      case 'last_month':
        start = formatDate(new Date(now.getFullYear(), now.getMonth() - 1, 1));
        end = formatDate(new Date(now.getFullYear(), now.getMonth(), 0));
        break;
      case 'this_quarter':
        const currentQuarter = Math.floor(now.getMonth() / 3);
        start = formatDate(new Date(now.getFullYear(), currentQuarter * 3, 1));
        end = formatDate(new Date(now.getFullYear(), currentQuarter * 3 + 3, 0));
        break;
      case 'last_quarter':
        const lastQuarter = Math.floor(now.getMonth() / 3) - 1;
        start = formatDate(new Date(now.getFullYear(), lastQuarter * 3, 1));
        end = formatDate(new Date(now.getFullYear(), lastQuarter * 3 + 3, 0));
        break;
      case 'this_year':
        start = formatDate(new Date(now.getFullYear(), 0, 1));
        end = formatDate(new Date(now.getFullYear(), 11, 31));
        break;
      case 'last_year':
        start = formatDate(new Date(now.getFullYear() - 1, 0, 1));
        end = formatDate(new Date(now.getFullYear() - 1, 11, 31));
        break;
      case 'custom':
      default:
        return;
    }

    setStartDate(start);
    setEndDate(end);
  };

  const processedData = useMemo(() => {
    // Process Data_TVKT
    const tvktHeaders = data.dataTvkt[0] || [];
    const tvktRows = data.dataTvkt.slice(1);
    
    const nguoiThucHienIdx = findColumnIndex(tvktHeaders, ['nguoi_thuc_hien', 'người thực hiện', 'nguoi thuc hien', 'nhân viên', 'tư vấn', 'người tư vấn', 'chuyên viên', 'người']);
    let dateIdxTvkt = findColumnIndex(tvktHeaders, ['ngay_tv', 'ngày_tv', 'ngay tv', 'ngày tv']);
    if (dateIdxTvkt === -1) dateIdxTvkt = 2; // Fallback to Column C
    
    // Process Ban_ve_TVKT
    const banVeHeaders = data.banVeTvkt[0] || [];
    const banVeRows = data.banVeTvkt.slice(1);
    
    const bvPersonIdx = findColumnIndex(banVeHeaders, ['người', 'nhân viên', 'nguoi_thuc_hien', 'phụ trách', 'nguoi']);
    let dateIdxBanVe = findColumnIndex(banVeHeaders, ['ngay_nhan', 'ngày_nhận', 'ngay nhan', 'ngày nhận']);
    if (dateIdxBanVe === -1) dateIdxBanVe = 2; // Fallback to Column C
    const statusIdx = findColumnIndex(banVeHeaders, ['tbsx']);

    const unitsSet = new Set<string>();

    const parseDate = (dateVal: any) => {
      if (!dateVal) return null;
      if (typeof dateVal === 'number') {
        // Google Sheets serial date (days since Dec 30, 1899)
        return new Date((dateVal - 25569) * 86400 * 1000).getTime();
      }
      const dateStr = dateVal.toString().trim();
      const parts = dateStr.split(/[-/]/);
      if (parts.length === 3) {
        if (parts[0].length === 4) return new Date(`${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}T00:00:00`).getTime();
        return new Date(`${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}T00:00:00`).getTime();
      }
      const parsed = new Date(dateStr).getTime();
      return isNaN(parsed) ? null : parsed;
    };

    const startTimestamp = startDate ? new Date(`${startDate}T00:00:00`).getTime() : null;
    const endTimestamp = endDate ? new Date(`${endDate}T23:59:59.999`).getTime() : null;

    const isRowInDateRange = (row: any[], dateIdx: number) => {
      if (dateIdx < 0 || (!startTimestamp && !endTimestamp)) return true;
      const rowDateVal = row[dateIdx];
      if (!rowDateVal) return true;
      const rowTime = parseDate(rowDateVal);
      if (!rowTime || isNaN(rowTime)) return true;
      
      if (startTimestamp && rowTime < startTimestamp) return false;
      if (endTimestamp && rowTime > endTimestamp) return false;
      return true;
    };

    const isRowInUnit = (row: any[], personIdx: number) => {
      if (selectedUnit === 'All') return true;
      if (personIdx < 0) return false;
      const cellValue = row[personIdx];
      const people = splitNames(cellValue);
      return people.some(p => extractUnit(p) === selectedUnit);
    };

    // Extract all units first to populate the filter dropdown
    [...tvktRows, ...banVeRows].forEach(row => {
      const tvktPeople = nguoiThucHienIdx >= 0 ? splitNames(row[nguoiThucHienIdx]) : [];
      const bvPeople = bvPersonIdx >= 0 ? splitNames(row[bvPersonIdx]) : [];
      [...tvktPeople, ...bvPeople].forEach(p => unitsSet.add(extractUnit(p)));
    });

    const filteredTvktRows = tvktRows.filter(row => isRowInDateRange(row, dateIdxTvkt) && isRowInUnit(row, nguoiThucHienIdx));
    const tuVanStats = filteredTvktRows.reduce((acc: Record<string, number>, row) => {
      const cellValue = nguoiThucHienIdx >= 0 ? row[nguoiThucHienIdx] : '';
      const people = splitNames(cellValue);
      
      if (people.length > 0) {
        people.forEach(person => {
          const unit = extractUnit(person);
          if (selectedUnit === 'All' || unit === selectedUnit) {
            acc[person] = (acc[person] || 0) + 1;
          }
        });
      } else {
        if (selectedUnit === 'All') {
          acc['Unknown'] = (acc['Unknown'] || 0) + 1;
        }
      }
      return acc;
    }, {});

    const tuVanChartData = Object.entries(tuVanStats)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);
      
    const totalTuVan = tuVanChartData.reduce((sum, item) => sum + item.count, 0);

    const filteredBanVeRows = banVeRows.filter(row => isRowInDateRange(row, dateIdxBanVe) && isRowInUnit(row, bvPersonIdx));
    const totalBanVe = filteredBanVeRows.length;
    
    const deployedBanVe = statusIdx >= 0 
      ? filteredBanVeRows.filter(row => {
          const status = row[statusIdx]?.toString().trim() || '';
          return status !== '';
        }).length
      : totalBanVe;

    const banVePersonStats = filteredBanVeRows.reduce((acc: Record<string, number>, row) => {
      const cellValue = bvPersonIdx >= 0 ? row[bvPersonIdx] : '';
      const people = splitNames(cellValue);
      
      if (people.length > 0) {
        people.forEach(person => {
          const unit = extractUnit(person);
          if (selectedUnit === 'All' || unit === selectedUnit) {
            acc[person] = (acc[person] || 0) + 1;
          }
        });
      } else {
        if (selectedUnit === 'All') {
          acc['Unknown'] = (acc['Unknown'] || 0) + 1;
        }
      }
      return acc;
    }, {});

    const banVePersonChartData = Object.entries(banVePersonStats)
      .map(([name, count]) => ({ name, count }))
      .sort((a, b) => b.count - a.count);

    return {
      tuVanChartData,
      totalTuVan,
      banVePersonChartData,
      totalBanVe,
      deployedBanVe,
      statusIdx,
      availableUnits: Array.from(unitsSet).sort()
    };
  }, [data, startDate, endDate, selectedUnit]);

  const { tuVanChartData, totalTuVan, banVePersonChartData, totalBanVe, deployedBanVe, statusIdx, availableUnits } = processedData;

  return (
    <div className="space-y-8">
      {/* Filters */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div className="flex items-center gap-2 mb-4">
          <Filter className="w-5 h-5 text-indigo-600" />
          <h3 className="text-lg font-bold text-slate-900">Bộ lọc dữ liệu</h3>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Thời gian</label>
            <select 
              value={quickDate}
              onChange={(e) => handleQuickDateChange(e.target.value)}
              className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
            >
              <option value="all_time">Tất cả (Từ trước đến nay)</option>
              <option value="this_month">Tháng này</option>
              <option value="last_month">Tháng trước</option>
              <option value="this_quarter">Quý này</option>
              <option value="last_quarter">Quý trước</option>
              <option value="this_year">Năm nay</option>
              <option value="last_year">Năm ngoái</option>
              <option value="custom">Tùy chỉnh</option>
            </select>
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Từ ngày</label>
            <input 
              type="date" 
              value={startDate}
              onChange={(e) => { setStartDate(e.target.value); setQuickDate('custom'); }}
              className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Đến ngày</label>
            <input 
              type="date" 
              value={endDate}
              onChange={(e) => { setEndDate(e.target.value); setQuickDate('custom'); }}
              className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-slate-700 mb-1">Đơn vị</label>
            <select 
              value={selectedUnit}
              onChange={(e) => setSelectedUnit(e.target.value)}
              className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500"
            >
              <option value="All">Tất cả đơn vị</option>
              {availableUnits.map(unit => (
                <option key={unit} value={unit}>{unit}</option>
              ))}
            </select>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
        {/* Left Column: Consultations */}
        <div className="space-y-6">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <h3 className="text-slate-500 font-medium text-sm mb-1">Tổng số đi tư vấn</h3>
            <div className="text-3xl font-bold text-amber-600">{totalTuVan}</div>
          </div>

          {/* Chart 1: Số tư vấn theo người thực hiện */}
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col">
            <h3 className="text-lg font-bold text-slate-900 mb-6">Số tư vấn theo người thực hiện (Top {topNTuVan > 100 ? 'Tất cả' : topNTuVan})</h3>
            {tuVanChartData.length > 0 ? (
              <div className="h-80 flex-grow">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={tuVanChartData.slice(0, topNTuVan)} layout="vertical" margin={{ top: 5, right: 30, left: 0, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
                    <XAxis type="number" />
                    <YAxis dataKey="name" type="category" width={160} tick={<CustomYAxisTick />} />
                    <Tooltip 
                      contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                      cursor={{ fill: '#f1f5f9' }}
                    />
                    <Bar dataKey="count" fill="#10b981" radius={[0, 4, 4, 0]} name="Số lượng tư vấn" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            ) : (
              <div className="h-80 flex items-center justify-center text-slate-400">
                Không tìm thấy cột "người thực hiện" hoặc không có dữ liệu
              </div>
            )}
            <div className="mt-4 pt-4 border-t border-slate-100 flex items-center justify-end gap-2">
              <label className="text-sm font-medium text-slate-600">Hiển thị:</label>
              <select 
                value={topNTuVan}
                onChange={(e) => setTopNTuVan(Number(e.target.value))}
                className="px-3 py-1.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50"
              >
                <option value={5}>Top 5</option>
                <option value={10}>Top 10</option>
                <option value={15}>Top 15</option>
                <option value={20}>Top 20</option>
                <option value={50}>Top 50</option>
                <option value={1000}>Tất cả</option>
              </select>
            </div>
          </div>
        </div>

        {/* Right Column: Drawings */}
        <div className="space-y-6">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
            <div className="flex justify-between items-start">
              <div>
                <h3 className="text-slate-500 font-medium text-sm mb-1">Bản vẽ triển khai / Tổng số</h3>
                <div className="text-3xl font-bold text-slate-900">
                  <span className="text-emerald-600">{deployedBanVe}</span>
                  <span className="text-slate-300 mx-2">/</span>
                  <span>{totalBanVe}</span>
                </div>
                {statusIdx < 0 && (
                  <p className="text-[10px] text-slate-400 mt-1">*(Không tìm thấy cột TBSX)*</p>
                )}
              </div>
              <div className="text-right">
                <h3 className="text-slate-500 font-medium text-sm mb-1">Tỷ lệ</h3>
                <div className="text-3xl font-bold text-indigo-600">
                  {totalBanVe > 0 ? Math.round((deployedBanVe / totalBanVe) * 100) : 0}%
                </div>
              </div>
            </div>
          </div>

          {/* Chart 2: Số bản vẽ cho từng người */}
          <div className="bg-white rounded-2xl shadow-sm border border-slate-200 p-6 flex flex-col">
            <h3 className="text-lg font-bold text-slate-900 mb-6">Số bản vẽ cho từng người (Top {topNBanVe > 100 ? 'Tất cả' : topNBanVe})</h3>
            {banVePersonChartData.length > 0 ? (
              <div className="h-80 flex-grow">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={banVePersonChartData.slice(0, topNBanVe)} layout="vertical" margin={{ top: 5, right: 30, left: 0, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e2e8f0" />
                    <XAxis type="number" />
                    <YAxis dataKey="name" type="category" width={160} tick={<CustomYAxisTick />} />
                    <Tooltip 
                      contentStyle={{ borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
                      cursor={{ fill: '#f1f5f9' }}
                    />
                    <Bar dataKey="count" fill="#4f46e5" radius={[0, 4, 4, 0]} name="Số bản vẽ" />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            ) : (
              <div className="h-80 flex items-center justify-center text-slate-400">
                Không có dữ liệu
              </div>
            )}
            <div className="mt-4 pt-4 border-t border-slate-100 flex items-center justify-end gap-2">
              <label className="text-sm font-medium text-slate-600">Hiển thị:</label>
              <select 
                value={topNBanVe}
                onChange={(e) => setTopNBanVe(Number(e.target.value))}
                className="px-3 py-1.5 text-sm border border-slate-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50"
              >
                <option value={5}>Top 5</option>
                <option value={10}>Top 10</option>
                <option value={15}>Top 15</option>
                <option value={20}>Top 20</option>
                <option value={50}>Top 50</option>
                <option value={1000}>Tất cả</option>
              </select>
            </div>
          </div>
        </div>
      </div>

      {/* Detailed Stats Table */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 bg-slate-50">
          <h3 className="text-lg font-bold text-slate-900">Thống kê chi tiết: Số lượng tư vấn theo người</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="border-b border-slate-200 text-sm text-slate-500">
                <th className="px-6 py-3 font-medium">Người thực hiện</th>
                <th className="px-6 py-3 font-medium">Số lượng tư vấn</th>
                <th className="px-6 py-3 font-medium">Tỷ lệ (%)</th>
              </tr>
            </thead>
            <tbody className="text-sm text-slate-700">
              {tuVanChartData.map((row, idx) => (
                <tr key={idx} className="border-b border-slate-100 hover:bg-slate-50">
                  <td className="px-6 py-3 font-medium text-slate-900">{row.name}</td>
                  <td className="px-6 py-3">{row.count}</td>
                  <td className="px-6 py-3">
                    {totalTuVan > 0 ? Math.round((row.count / totalTuVan) * 100) : 0}%
                  </td>
                </tr>
              ))}
              {tuVanChartData.length === 0 && (
                <tr>
                  <td colSpan={3} className="px-6 py-8 text-center text-slate-400">
                    Không có dữ liệu
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
