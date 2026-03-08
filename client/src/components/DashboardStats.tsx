import React from 'react';
import { Chart as ChartJS, ArcElement, Tooltip, Legend, CategoryScale, LinearScale, BarElement } from 'chart.js';
import { Doughnut, Bar } from 'react-chartjs-2';
import { TrendingUp, Calendar, CheckCircle, Clock } from 'lucide-react';

ChartJS.register(ArcElement, Tooltip, Legend, CategoryScale, LinearScale, BarElement);

interface DashboardStatsProps {
  meetings: any[];
  items: any[];
  period: string;
  selectedMeeting?: string;
}

const DashboardStats: React.FC<DashboardStatsProps> = ({ meetings, items, period, selectedMeeting = 'all' }) => {
  // Filter items by selected meeting if not 'all'
  // When viewing live stats for a specific meeting, show only items from that meeting
  // to match the meeting header statistics exactly
  
  const filteredItems = selectedMeeting === 'all' 
    ? items 
    : items.filter(item => {
        // Extract date part from item.meetingDate for comparison
        // item.meetingDate is like "2025-07-15T10:00:00.000Z"
        // selectedMeeting is like "2025-07-15"
        const itemDateOnly = item.meetingDate.split('T')[0];
        return itemDateOnly === selectedMeeting;
      });
  
  // Filter meetings by selected meeting if not 'all'
  const filteredMeetings = selectedMeeting === 'all' 
    ? meetings 
    : meetings.filter(meeting => meeting.meetingDate === selectedMeeting);
  
  // Calculate stats from filtered data
  // For category counts, use the same logic as filtered items to include unclosed previous items
  const safetyCount = filteredItems.filter(item => item.type === 'Safety Ideas').length;
  const businessCount = filteredItems.filter(item => item.type === 'Business Ideas').length;
  const nearMissCount = filteredItems.filter(item => item.type === 'Near Miss').length;
  
  const submittedCount = filteredItems.filter(item => item.status === 'Submitted').length;
  const discussionCount = filteredItems.filter(item => item.status === 'In Discussion').length;
  const actionedCount = filteredItems.filter(item => item.status === 'Actions' || item.status === 'Actioned').length;
  const closedCount = filteredItems.filter(item => item.status === 'Closed').length;

  // Category breakdown data - Match the colors from status badges
  const categoryData = {
    labels: ['Safety Ideas', 'Business Ideas', 'Near Miss'],
    datasets: [{
      data: [safetyCount, businessCount, nearMissCount],
      backgroundColor: ['#EF4444', '#3B82F6', '#F97316'],  // Red-600, Blue-600, Orange-600 matching card colors
      borderColor: '#ffffff',
      borderWidth: 2,
    }]
  };

  // Status distribution data - Match the colors from stat cards above
  const statusData = {
    labels: ['Submitted', 'In Discussion', 'Being Actioned', 'Closed'],
    datasets: [{
      label: 'Items',
      data: [submittedCount, discussionCount, actionedCount, closedCount],
      backgroundColor: ['#3B82F6', '#EAB308', '#A855F7', '#10B981'],  // Blue-600, Yellow-600, Purple-600, Green-600
      borderColor: '#ffffff',
      borderWidth: 2,
    }]
  };

  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    layout: {
      padding: {
        bottom: 10
      }
    },
    plugins: {
      legend: {
        position: 'bottom' as const,
        align: 'center' as const,
        labels: {
          padding: 10,
          boxWidth: 15,
          boxHeight: 15,
          font: {
            size: 11,
            weight: 'normal' as const
          },
          usePointStyle: true,
          pointStyle: 'circle' as const,
          generateLabels: (chart: any) => {
            const data = chart.data;
            if (data.labels.length && data.datasets.length) {
              return data.labels.map((label: string, i: number) => {
                const value = data.datasets[0].data[i];
                return {
                  text: `${label}: ${value}`,
                  fillStyle: data.datasets[0].backgroundColor[i],
                  strokeStyle: data.datasets[0].backgroundColor[i],
                  lineWidth: 0,
                  hidden: false,
                  index: i
                };
              });
            }
            return [];
          }
        }
      },
      tooltip: {
        backgroundColor: 'rgba(0, 0, 0, 0.9)',
        padding: 12,
        cornerRadius: 6,
        titleFont: {
          size: 13,
          weight: 'bold' as const
        },
        bodyFont: {
          size: 12
        }
      }
    }
  };

  const barOptions = {
    ...chartOptions,
    plugins: {
      ...chartOptions.plugins,
      legend: {
        display: false
      }
    },
    scales: {
      y: {
        beginAtZero: true,
        ticks: {
          stepSize: 5,
          font: {
            size: 11
          },
          color: '#6B7280'
        },
        grid: {
          color: '#E5E7EB',
          drawBorder: false
        }
      },
      x: {
        ticks: {
          font: {
            size: 11,
            weight: 'normal' as const
          },
          color: '#374151'
        },
        grid: {
          display: false,
          drawBorder: false
        }
      }
    }
  };

  const periodText = period === 'all' ? 'All Time' : 
                    period === 'relevant' ? 'Relevant Items' :
                    period === 'q1' ? 'Q1 2025' :
                    period === 'q2' ? 'Q2 2025' :
                    period === 'q3' ? 'Q3 2025' :
                    period === 'q4' ? 'Q4 2025' :
                    'Last 6 Months';

  return (
    <div className="space-y-4">
      {/* Live Stats Header */}
      {selectedMeeting !== 'all' && (
        <div className="flex items-center justify-center gap-2 py-2 px-4 bg-gradient-to-r from-green-50 to-emerald-50 border border-green-200 rounded-lg">
          <div className="w-2.5 h-2.5 bg-green-500 rounded-full animate-pulse"></div>
          <span className="text-sm font-semibold text-green-700">Live Meeting Statistics</span>
        </div>
      )}
      
      {/* Enhanced Stats Row */}
      <div className={`grid gap-3 ${selectedMeeting === 'all' ? 'grid-cols-2 md:grid-cols-6' : 'grid-cols-2 md:grid-cols-5'}`}>
        {/* Only show Meeting count when viewing all meetings */}
        {selectedMeeting === 'all' && (
          <div className="bg-white rounded-lg border border-gray-200 p-3 text-center shadow-sm">
            <div className="text-xl font-bold text-gray-900">{filteredMeetings.length}</div>
            <div className="text-xs text-gray-600 mt-1">Meetings</div>
          </div>
        )}
        
        <div className="bg-gradient-to-br from-blue-50 to-indigo-50 rounded-lg border border-blue-200 p-3 text-center shadow-sm">
          <div className="text-xl font-bold text-blue-900">{filteredItems.length}</div>
          <div className="text-xs text-blue-700 mt-1 font-medium">Total Items</div>
        </div>
        
        <div className="bg-gradient-to-br from-blue-50 to-blue-100 rounded-lg border border-blue-300 p-3 text-center shadow-sm">
          <div className="text-2xl font-bold text-blue-700">{submittedCount}</div>
          <div className="text-xs text-blue-600 mt-1 font-medium">Submitted</div>
        </div>
        
        <div className="bg-gradient-to-br from-amber-50 to-yellow-100 rounded-lg border border-amber-300 p-3 text-center shadow-sm">
          <div className="text-2xl font-bold text-amber-700">{discussionCount}</div>
          <div className="text-xs text-amber-600 mt-1 font-medium">Discussions</div>
        </div>
        
        <div className="bg-gradient-to-br from-purple-50 to-violet-100 rounded-lg border border-purple-300 p-3 text-center shadow-sm">
          <div className="text-2xl font-bold text-purple-700">{actionedCount}</div>
          <div className="text-xs text-purple-600 mt-1 font-medium">Actions</div>
        </div>
        
        <div className="bg-gradient-to-br from-emerald-50 to-green-100 rounded-lg border border-emerald-300 p-3 text-center shadow-sm">
          <div className="text-2xl font-bold text-emerald-700">{closedCount}</div>
          <div className="text-xs text-emerald-600 mt-1 font-medium">Closed</div>
        </div>
      </div>

      {/* Simplified Charts */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* Category Distribution */}
        <div className="bg-white rounded-lg border border-gray-200 p-4">
          <div className="flex items-center gap-2 mb-3">
            <h4 className="text-sm font-medium text-gray-700">Item Categories by Type</h4>
            {selectedMeeting !== 'all' && (
              <div className="flex items-center gap-1 text-xs text-green-600">
                <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse"></div>
                Live
              </div>
            )}
          </div>
          <div style={{ height: '200px', position: 'relative' }}>
            <Doughnut data={categoryData} options={chartOptions} />
          </div>
        </div>

        {/* Status Distribution */}
        <div className="bg-white rounded-lg border border-gray-200 p-4">
          <div className="flex items-center gap-2 mb-3">
            <h4 className="text-sm font-medium text-gray-700">Action Status Distribution</h4>
            {selectedMeeting !== 'all' && (
              <div className="flex items-center gap-1 text-xs text-green-600">
                <div className="w-1.5 h-1.5 bg-green-500 rounded-full animate-pulse"></div>
                Live
              </div>
            )}
          </div>
          <div style={{ height: '200px' }}>
            <Bar data={statusData} options={barOptions} />
          </div>
        </div>
      </div>
    </div>
  );
};

export default DashboardStats;