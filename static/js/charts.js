// Admin Analytics Charts with Chart.js
document.addEventListener('DOMContentLoaded', () => {
    const typeCanvas = document.getElementById('leaveTypeChart');
    const monthCanvas = document.getElementById('monthlyTrendChart');

    if (!typeCanvas && !monthCanvas) return;

    fetch('/api/chart-data')
        .then(res => res.json())
        .then(data => {
            // 1. Leave Type Donut Chart
            if (typeCanvas) {
                new Chart(typeCanvas.getContext('2d'), {
                    type: 'doughnut',
                    data: {
                        labels: data.type_labels,
                        datasets: [{
                            data: data.type_data,
                            backgroundColor: data.type_colors,
                            borderWidth: 2,
                            borderColor: '#ffffff',
                            hoverOffset: 6
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: {
                                position: 'bottom',
                                labels: {
                                    boxWidth: 12,
                                    padding: 16,
                                    font: { family: "'Plus Jakarta Sans', sans-serif", size: 12, weight: '500' }
                                }
                            },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return ` ${context.label}: ${context.raw} day(s) taken`;
                                    }
                                }
                            }
                        },
                        cutout: '68%'
                    }
                });
            }

            // 2. Monthly Trend Bar Chart
            if (monthCanvas) {
                new Chart(monthCanvas.getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: data.month_labels,
                        datasets: [{
                            label: 'Approved Leave Days',
                            data: data.month_data,
                            backgroundColor: 'rgba(79, 70, 229, 0.85)',
                            hoverBackgroundColor: 'rgba(67, 56, 202, 1)',
                            borderRadius: 6,
                            borderSkipped: false
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        scales: {
                            y: {
                                beginAtZero: true,
                                grid: { color: 'rgba(241, 245, 249, 1)' },
                                ticks: { stepSize: 2, font: { family: "'Plus Jakarta Sans', sans-serif", size: 11 } }
                            },
                            x: {
                                grid: { display: false },
                                ticks: { font: { family: "'Plus Jakarta Sans', sans-serif", size: 11 } }
                            }
                        },
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return ` ${context.raw} days taken`;
                                    }
                                }
                            }
                        }
                    }
                });
            }
        })
        .catch(err => console.error('Failed to load chart analytics:', err));
});
