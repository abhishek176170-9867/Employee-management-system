// Global interactive scripts for ELMS
document.addEventListener('DOMContentLoaded', () => {
    // 1. Auto-dismiss alerts after 5 seconds
    const alerts = document.querySelectorAll('.alert-banner');
    alerts.forEach(alert => {
        setTimeout(() => {
            alert.style.transition = 'opacity 0.5s ease-out, transform 0.5s ease-out';
            alert.style.opacity = '0';
            alert.style.transform = 'translateY(-10px)';
            setTimeout(() => alert.remove(), 500);
        }, 5000);
    });

    // 2. Leave Application Date Calculator
    const startDateInput = document.getElementById('start_date');
    const endDateInput = document.getElementById('end_date');
    const halfDayCheck = document.getElementById('is_half_day');
    const leaveTypeSelect = document.getElementById('leave_type_id');
    const daysCalcBadge = document.getElementById('calculated_days_badge');
    const daysCalcText = document.getElementById('calculated_days_text');
    const balanceWarning = document.getElementById('balance_warning');
    const submitBtn = document.getElementById('submit_leave_btn');

    function checkLiveDuration() {
        if (!startDateInput || !endDateInput) return;

        const isHalfDay = halfDayCheck ? halfDayCheck.checked : false;
        
        if (isHalfDay && startDateInput.value) {
            endDateInput.value = startDateInput.value;
            endDateInput.readOnly = true;
            endDateInput.classList.add('bg-slate-100', 'cursor-not-allowed');
            endDateInput.removeAttribute('disabled');
        } else if (endDateInput) {
            endDateInput.readOnly = false;
            endDateInput.classList.remove('bg-slate-100', 'cursor-not-allowed');
            endDateInput.removeAttribute('disabled');
        }

        const startVal = startDateInput.value;
        const endVal = endDateInput.value;

        if (!startVal || !endVal) {
            if (daysCalcBadge) daysCalcBadge.classList.add('hidden');
            if (balanceWarning) balanceWarning.classList.add('hidden');
            return;
        }

        fetch('/api/calculate-days', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                start_date: startVal,
                end_date: endVal,
                is_half_day: isHalfDay
            })
        })
        .then(res => res.json())
        .then(data => {
            if (daysCalcBadge && daysCalcText) {
                daysCalcBadge.classList.remove('hidden');
                if (data.valid) {
                    daysCalcText.textContent = `${data.days} working day(s) calculated (excluding weekends & holidays)`;
                    daysCalcBadge.className = 'mt-2 p-3 rounded-xl text-sm font-medium bg-emerald-50 text-emerald-700 border border-emerald-200 flex items-center gap-2';
                    
                    // Check balance availability if leave type is selected
                    if (leaveTypeSelect && balanceWarning) {
                        const selectedOption = leaveTypeSelect.options[leaveTypeSelect.selectedIndex];
                        const available = parseFloat(selectedOption.getAttribute('data-available') || '0');
                        if (data.days > available) {
                            balanceWarning.classList.remove('hidden');
                            balanceWarning.textContent = `Warning: Requested ${data.days} days exceeds your available balance (${available} days). Application may be rejected.`;
                        } else {
                            balanceWarning.classList.add('hidden');
                        }
                    }
                } else {
                    daysCalcText.textContent = data.message;
                    daysCalcBadge.className = 'mt-2 p-3 rounded-xl text-sm font-medium bg-amber-50 text-amber-700 border border-amber-200 flex items-center gap-2';
                    if (balanceWarning) balanceWarning.classList.add('hidden');
                }
            }
        })
        .catch(err => console.error('Calculation error:', err));
    }

    if (startDateInput && endDateInput) {
        startDateInput.addEventListener('change', () => {
            if (halfDayCheck && halfDayCheck.checked) {
                endDateInput.value = startDateInput.value;
            }
            checkLiveDuration();
        });
        endDateInput.addEventListener('change', checkLiveDuration);
        if (halfDayCheck) {
            halfDayCheck.addEventListener('change', () => {
                if (halfDayCheck.checked && startDateInput.value) {
                    endDateInput.value = startDateInput.value;
                }
                checkLiveDuration();
            });
        }
        if (leaveTypeSelect) {
            leaveTypeSelect.addEventListener('change', checkLiveDuration);
        }
    }
});

// Modal helpers
function openModal(id) {
    const el = document.getElementById(id);
    if (el) {
        el.classList.remove('hidden');
        el.classList.add('flex');
    }
}

function closeModal(id) {
    const el = document.getElementById(id);
    if (el) {
        el.classList.add('hidden');
        el.classList.remove('flex');
    }
}
