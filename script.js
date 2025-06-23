// URL Apps Script Web App
const SCRIPT_URL =
  'https://script.google.com/macros/s/AKfycbx7tVhnmIIhVOz7fZ-4b0UWZZE2F6HmtRqI8_A0afz5HdXXJ-OVuv9Fim9Qc0sHDvEn9Q/exec';

// DOM Elements
const loginSection = document.getElementById('loginSection');
const dashboardSection = document.getElementById('dashboardSection');
const siswaDashboard = document.getElementById('siswaDashboard');
const guruDashboard = document.getElementById('guruDashboard');
const userTypeSelect = document.getElementById('userType');
const siswaFields = document.getElementById('siswaFields');
const guruFields = document.getElementById('guruFields');
const loginBtn = document.getElementById('loginBtn');
const logoutBtn = document.getElementById('logoutBtn');
const welcomeMessage = document.getElementById('welcomeMessage');
const userRole = document.getElementById('userRole');
const currentDate = document.getElementById('currentDate');
const hadirBtn = document.getElementById('hadirBtn');
const izinBtn = document.getElementById('izinBtn');
const absenBtn = document.getElementById('absenBtn');
const riwayatBody = document.getElementById('riwayatBody');
const filterKelas = document.getElementById('filterKelas');
const filterTanggal = document.getElementById('filterTanggal');
const filterBtn = document.getElementById('filterBtn');
const exportBtn = document.getElementById('exportBtn');
const kehadiranBody = document.getElementById('kehadiranBody');
const newKelas = document.getElementById('newKelas');
const tambahSiswaBtn = document.getElementById('tambahSiswaBtn');
const tambahKelasBtn = document.getElementById('tambahKelasBtn');
const loginAlert = document.getElementById('loginAlert');
const absensiAlert = document.getElementById('absensiAlert');
const guruAlert = document.getElementById('guruAlert');
const tabs = document.querySelectorAll('.tab');
const tabContents = document.querySelectorAll('.tab-content');

// Global Variables
let currentUser = null;

// Event Listeners
document.addEventListener('DOMContentLoaded', () => {
  // Set current date
  const today = new Date();
  if (currentDate) {
    currentDate.textContent = today.toLocaleDateString('id-ID', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
  }

  // Set default date for filter
  if (filterTanggal) {
    filterTanggal.valueAsDate = today;
  }

  // Check if user is already logged in
  checkLoginStatus();
});

// Event Listeners with null checks
if (userTypeSelect) {
  userTypeSelect.addEventListener('change', (e) => {
    if (e.target.value === 'siswa') {
      if (siswaFields) siswaFields.classList.remove('hidden');
      if (guruFields) guruFields.classList.add('hidden');
    } else {
      if (siswaFields) siswaFields.classList.add('hidden');
      if (guruFields) guruFields.classList.remove('hidden');
    }
  });
}

if (loginBtn) loginBtn.addEventListener('click', login);
if (logoutBtn) logoutBtn.addEventListener('click', logout);
if (hadirBtn) hadirBtn.addEventListener('click', () => submitAbsensi('Hadir'));
if (izinBtn) izinBtn.addEventListener('click', () => submitAbsensi('Izin'));
if (absenBtn) absenBtn.addEventListener('click', () => submitAbsensi('Absen'));
if (filterBtn) filterBtn.addEventListener('click', filterKehadiran);
if (exportBtn) exportBtn.addEventListener('click', exportToExcel);
if (tambahSiswaBtn) tambahSiswaBtn.addEventListener('click', tambahSiswa);
if (tambahKelasBtn) tambahKelasBtn.addEventListener('click', tambahKelas);

if (tabs) {
  tabs.forEach((tab) => {
    tab.addEventListener('click', () => {
      tabs.forEach((t) => t.classList.remove('active'));
      tabContents.forEach((c) => c.classList.remove('active'));

      tab.classList.add('active');
      const targetTab = document.getElementById(`${tab.dataset.tab}Tab`);
      if (targetTab) targetTab.classList.add('active');
    });
  });
}

// Functions
function checkLoginStatus() {
  try {
    const userData = localStorage.getItem('absensiUser');
    if (userData) {
      currentUser = JSON.parse(userData);
      showDashboard();
    }
  } catch (error) {
    console.error('Error checking login status:', error);
    localStorage.removeItem('absensiUser');
  }
}

function login() {
  const userType = userTypeSelect ? userTypeSelect.value : 'siswa';
  let nis = '',
    username = '',
    password = '';

  if (userType === 'siswa') {
    const nisField = document.getElementById('nis');
    const passwordField = document.getElementById('passwordSiswa');
    if (nisField) nis = nisField.value.trim();
    if (passwordField) password = passwordField.value.trim();
  } else {
    const usernameField = document.getElementById('username');
    const passwordField = document.getElementById('passwordGuru');
    if (usernameField) username = usernameField.value.trim();
    if (passwordField) password = passwordField.value.trim();
  }

  // Validation
  if (
    (userType === 'siswa' && (!nis || !password)) ||
    (userType === 'guru' && (!username || !password))
  ) {
    showAlert(loginAlert, 'danger', 'Harap isi semua field');
    return;
  }

  // Show loading state
  if (loginBtn) {
    loginBtn.disabled = true;
    loginBtn.textContent = 'Loading...';
  }

  const params = new URLSearchParams();
  params.append('action', 'login');
  params.append('userType', userType);
  params.append('nis', nis);
  params.append('username', username);
  params.append('password', password);

  fetch(`${SCRIPT_URL}?${params.toString()}`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
  })
    .then((response) => {
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.json();
    })
    .then((data) => {
      if (data.success) {
        currentUser = data;
        localStorage.setItem('absensiUser', JSON.stringify(data));
        showDashboard();
        showAlert(loginAlert, 'success', 'Login berhasil!');
      } else {
        showAlert(loginAlert, 'danger', data.message || 'Login gagal');
      }
    })
    .catch((error) => {
      console.error('Login error:', error);
      showAlert(loginAlert, 'danger', 'Terjadi kesalahan. Silakan coba lagi.');
    })
    .finally(() => {
      if (loginBtn) {
        loginBtn.disabled = false;
        loginBtn.textContent = 'Login';
      }
    });
}

function showDashboard() {
  if (loginSection) loginSection.classList.add('hidden');
  if (dashboardSection) dashboardSection.classList.remove('hidden');

  if (welcomeMessage) {
    welcomeMessage.textContent = `Selamat datang, ${
      currentUser.nama || currentUser.username
    }!`;
  }

  if (userRole) {
    userRole.textContent = `Anda login sebagai ${
      currentUser.userType === 'siswa' ? 'Siswa' : 'Guru'
    }`;
  }

  if (currentUser.userType === 'siswa') {
    if (siswaDashboard) siswaDashboard.classList.remove('hidden');
    if (guruDashboard) guruDashboard.classList.add('hidden');
    loadRiwayatKehadiran();
  } else {
    if (siswaDashboard) siswaDashboard.classList.add('hidden');
    if (guruDashboard) guruDashboard.classList.remove('hidden');
    loadKelasForFilter();
    loadKelasForAddSiswa();
    filterKehadiran();
  }
}

function logout() {
  try {
    localStorage.removeItem('absensiUser');
    currentUser = null;

    if (dashboardSection) dashboardSection.classList.add('hidden');
    if (loginSection) loginSection.classList.remove('hidden');

    // Reset form
    const fields = ['nis', 'passwordSiswa', 'username', 'passwordGuru'];
    fields.forEach((fieldId) => {
      const field = document.getElementById(fieldId);
      if (field) field.value = '';
    });

    if (loginAlert) loginAlert.classList.add('hidden');
  } catch (error) {
    console.error('Logout error:', error);
  }
}

function submitAbsensi(status) {
  if (!currentUser || !currentUser.nis) {
    showAlert(absensiAlert, 'danger', 'Data user tidak valid');
    return;
  }

  const params = new URLSearchParams();
  params.append('action', 'submitAbsensi');
  params.append('nis', currentUser.nis);
  params.append('nama', currentUser.nama);
  params.append('kelas', currentUser.kelas);
  params.append('status', status);

  // Disable buttons during submission
  const buttons = [hadirBtn, izinBtn, absenBtn];
  buttons.forEach((btn) => {
    if (btn) {
      btn.disabled = true;
      btn.textContent = 'Loading...';
    }
  });

  fetch(`${SCRIPT_URL}?${params.toString()}`, {
    method: 'POST',
  })
    .then((response) => {
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      return response.json();
    })
    .then((data) => {
      if (data.success) {
        showAlert(
          absensiAlert,
          'success',
          `Absensi ${status} berhasil dicatat`
        );
        loadRiwayatKehadiran();
      } else {
        showAlert(
          absensiAlert,
          'danger',
          data.message || 'Gagal mencatat absensi'
        );
      }
    })
    .catch((error) => {
      console.error('Absensi error:', error);
      showAlert(
        absensiAlert,
        'danger',
        'Terjadi kesalahan. Silakan coba lagi.'
      );
    })
    .finally(() => {
      // Re-enable buttons
      if (hadirBtn) {
        hadirBtn.disabled = false;
        hadirBtn.textContent = 'Hadir';
      }
      if (izinBtn) {
        izinBtn.disabled = false;
        izinBtn.textContent = 'Izin';
      }
      if (absenBtn) {
        absenBtn.disabled = false;
        absenBtn.textContent = 'Absen';
      }
    });
}

function loadRiwayatKehadiran() {
  if (!currentUser || !currentUser.kelas) return;

  const params = new URLSearchParams();
  params.append('action', 'getKehadiran');
  params.append('kelas', currentUser.kelas);
  params.append('tanggal', 'all');

  fetch(`${SCRIPT_URL}?${params.toString()}`)
    .then((response) => response.json())
    .then((data) => {
      if (data.success && riwayatBody) {
        riwayatBody.innerHTML = '';

        // Filter hanya data siswa ini
        const filteredData = data.data.filter(
          (item) => item.nis === currentUser.nis
        );

        if (filteredData.length === 0) {
          riwayatBody.innerHTML =
            '<tr><td colspan="2" style="text-align: center;">Belum ada riwayat kehadiran</td></tr>';
          return;
        }

        filteredData.forEach((item) => {
          const row = document.createElement('tr');

          const dateCell = document.createElement('td');
          const date = new Date(item.waktu);
          dateCell.textContent = date.toLocaleDateString('id-ID', {
            weekday: 'short',
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
          });

          const statusCell = document.createElement('td');
          statusCell.textContent = item.status;

          // Add status classes
          if (item.status === 'Hadir') {
            statusCell.classList.add('status-present');
          } else if (item.status === 'Izin') {
            statusCell.classList.add('status-permit');
          } else {
            statusCell.classList.add('status-absent');
          }

          row.appendChild(dateCell);
          row.appendChild(statusCell);
          riwayatBody.appendChild(row);
        });
      }
    })
    .catch((error) => {
      console.error('Load riwayat error:', error);
    });
}

function loadKelasForFilter() {
  fetch(`${SCRIPT_URL}?action=getKelas`)
    .then((response) => response.json())
    .then((data) => {
      if (data.success && filterKelas) {
        filterKelas.innerHTML = '<option value="all">Semua Kelas</option>';
        data.data.forEach((kelas) => {
          const option = document.createElement('option');
          option.value = kelas;
          option.textContent = kelas;
          filterKelas.appendChild(option);
        });
      }
    })
    .catch((error) => {
      console.error('Load kelas filter error:', error);
    });
}

function loadKelasForAddSiswa() {
  fetch(`${SCRIPT_URL}?action=getKelas`)
    .then((response) => response.json())
    .then((data) => {
      if (data.success && newKelas) {
        newKelas.innerHTML = '';
        data.data.forEach((kelas) => {
          const option = document.createElement('option');
          option.value = kelas;
          option.textContent = kelas;
          newKelas.appendChild(option);
        });
      }
    })
    .catch((error) => {
      console.error('Load kelas add siswa error:', error);
    });
}

function filterKehadiran() {
  const kelas = filterKelas.value;
  const tanggal = filterTanggal.value;

  if (!tanggal) {
    alert('Pilih tanggal terlebih dahulu');
    return;
  }

  // Format tanggal untuk Apps Script
  const dateParts = tanggal.split('-');
  const formattedDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;

  const params = new URLSearchParams();
  params.append('action', 'getKehadiranByKelas');
  params.append('kelas', kelas);
  params.append('tanggal', formattedDate);

  // Tampilkan loading
  kehadiranBody.innerHTML =
    '<tr><td colspan="7" style="text-align: center;">Memuat data...</td></tr>';

  fetch(`${SCRIPT_URL}?${params.toString()}`)
    .then((response) => response.json())
    .then((data) => {
      if (data.success) {
        kehadiranBody.innerHTML = '';

        if (data.data.length === 0) {
          kehadiranBody.innerHTML =
            '<tr><td colspan="7" style="text-align: center;">Tidak ada data siswa</td></tr>';
          return;
        }

        data.data.forEach((item, index) => {
          const row = document.createElement('tr');

          // No
          const noCell = document.createElement('td');
          noCell.textContent = index + 1;

          // NIS
          const nisCell = document.createElement('td');
          nisCell.textContent = item.nis;

          // Nama
          const namaCell = document.createElement('td');
          namaCell.textContent = item.nama;

          // Kelas
          const kelasCell = document.createElement('td');
          kelasCell.textContent = item.kelas;

          // Status dengan dropdown
          const statusCell = document.createElement('td');
          const statusSelect = document.createElement('select');
          statusSelect.className = 'status-select';
          statusSelect.setAttribute('data-nis', item.nis);

          const statusOptions = [
            { value: '', text: 'Belum Absen', class: 'status-not-set' },
            { value: 'Hadir', text: 'Hadir', class: 'status-present' },
            { value: 'Izin', text: 'Izin', class: 'status-permit' },
            { value: 'Sakit', text: 'Sakit', class: 'status-sick' },
            { value: 'Alpha', text: 'Alpha', class: 'status-absent' },
          ];

          statusOptions.forEach((option) => {
            const optionElement = document.createElement('option');
            optionElement.value = option.value;
            optionElement.textContent = option.text;
            if (option.value === item.status) {
              optionElement.selected = true;
            }
            statusSelect.appendChild(optionElement);
          });

          // Set class berdasarkan status
          statusSelect.className = `status-select ${getStatusClass(
            item.status
          )}`;

          // Event listener untuk perubahan status
          statusSelect.addEventListener('change', function () {
            this.className = `status-select ${getStatusClass(this.value)}`;
          });

          statusCell.appendChild(statusSelect);

          // Waktu
          const waktuCell = document.createElement('td');
          if (item.waktu && item.waktu !== '') {
            try {
              const date = new Date(item.waktu);
              waktuCell.textContent = date.toLocaleDateString('id-ID', {
                weekday: 'short',
                year: 'numeric',
                month: 'short',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
              });
            } catch (error) {
              waktuCell.textContent = '-';
            }
          } else {
            waktuCell.textContent = '-';
            waktuCell.style.color = '#999';
          }

          // Action buttons
          const actionCell = document.createElement('td');
          const saveBtn = document.createElement('button');
          saveBtn.textContent = 'Simpan';
          saveBtn.className = 'btn-save';
          saveBtn.onclick = () => {
            const selectedStatus = statusSelect.value;
            if (!selectedStatus) {
              alert('Pilih status kehadiran terlebih dahulu');
              return;
            }
            updateKehadiran(item.nis, selectedStatus, formattedDate);
          };

          actionCell.appendChild(saveBtn);

          row.appendChild(noCell);
          row.appendChild(nisCell);
          row.appendChild(namaCell);
          row.appendChild(kelasCell);
          row.appendChild(statusCell);
          row.appendChild(waktuCell);
          row.appendChild(actionCell);
          kehadiranBody.appendChild(row);
        });
      } else {
        kehadiranBody.innerHTML = `<tr><td colspan="7" style="text-align: center; color: red;">Error: ${data.message}</td></tr>`;
      }
    })
    .catch((error) => {
      console.error(error);
      kehadiranBody.innerHTML =
        '<tr><td colspan="7" style="text-align: center; color: red;">Gagal memuat data</td></tr>';
    });
}

function getStatusClass(status) {
  switch (status) {
    case 'Hadir':
      return 'status-present';
    case 'Izin':
      return 'status-permit';
    case 'Sakit':
      return 'status-sick';
    case 'Alpha':
      return 'status-absent';
    default:
      return 'status-not-set';
  }
}

function updateKehadiran(nis, status, tanggal) {
  const params = new URLSearchParams();
  params.append('action', 'updateKehadiran');
  params.append('nis', nis);
  params.append('status', status);
  params.append('tanggal', tanggal);

  fetch(SCRIPT_URL, {
    method: 'POST',
    body: params,
  })
    .then((response) => response.json())
    .then((data) => {
      if (data.success) {
        // Refresh data setelah update
        filterKehadiran();
        alert('Status kehadiran berhasil diperbarui');
      } else {
        alert('Gagal memperbarui status: ' + data.message);
      }
    })
    .catch((error) => {
      console.error(error);
      alert('Terjadi kesalahan saat memperbarui status');
    });
}

function updateAllKehadiran() {
  const selects = document.querySelectorAll('.status-select');
  const updates = [];

  selects.forEach((select) => {
    if (select.value) {
      updates.push({
        nis: select.getAttribute('data-nis'),
        status: select.value,
      });
    }
  });

  if (updates.length === 0) {
    alert('Tidak ada data untuk diperbarui');
    return;
  }

  const tanggal = filterTanggal.value;
  const dateParts = tanggal.split('-');
  const formattedDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;

  const params = new URLSearchParams();
  params.append('action', 'updateBulkKehadiran');
  params.append('updates', JSON.stringify(updates));
  params.append('tanggal', formattedDate);

  fetch(SCRIPT_URL, {
    method: 'POST',
    body: params,
  })
    .then((response) => response.json())
    .then((data) => {
      if (data.success) {
        filterKehadiran();
        alert(data.data);
      } else {
        alert('Gagal memperbarui status: ' + data.message);
      }
    })
    .catch((error) => {
      console.error(error);
      alert('Terjadi kesalahan saat memperbarui status');
    });
}

function exportToExcel() {
  const kelas = filterKelas.value;
  const tanggal = filterTanggal.value;
  const exportBtn = document.getElementById('exportBtn');

  // Validasi data
  if (!tanggal) {
    alert('Pilih tanggal terlebih dahulu');
    return;
  }

  // Tampilkan loading state
  const originalText = exportBtn.innerHTML;
  exportBtn.disabled = true;
  exportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Mengekspor...';

  try {
    // Ambil data dari tabel yang sudah ada di halaman
    const table = document.getElementById('kehadiranTable');
    const rows = table.querySelectorAll('tbody tr');

    // Validasi apakah ada data
    if (rows.length === 0) {
      alert('Tidak ada data untuk diekspor');
      return;
    }

    // Siapkan data untuk Excel
    const excelData = [];

    // Tambahkan header (tanpa kolom Aksi)
    const headers = ['No', 'NIS', 'Nama', 'Kelas', 'Status', 'Waktu'];
    excelData.push(headers);

    // Tambahkan data
    rows.forEach((row, index) => {
      const cells = row.querySelectorAll('td');

      // Skip jika row kosong atau hanya pesan error/loading
      if (cells.length < 6) return;

      const rowData = [];

      // No
      rowData.push(index + 1);

      // NIS
      rowData.push(cells[1].textContent.trim());

      // Nama
      rowData.push(cells[2].textContent.trim());

      // Kelas
      rowData.push(cells[3].textContent.trim());

      // Status - ambil dari dropdown/select value
      const statusCell = cells[4];
      const statusSelect = statusCell.querySelector('select.status-select');
      let statusValue = '';

      if (statusSelect) {
        statusValue = statusSelect.value;
        // Konversi ke text yang lebih readable
        switch (statusValue) {
          case '':
            statusValue = 'Belum Absen';
            break;
          case 'Hadir':
            statusValue = 'Hadir';
            break;
          case 'Izin':
            statusValue = 'Izin';
            break;
          case 'Sakit':
            statusValue = 'Sakit';
            break;
          case 'Alpha':
            statusValue = 'Alpha';
            break;
          default:
            statusValue = statusValue || 'Belum Absen';
        }
      } else {
        // Fallback jika tidak ada select (ambil dari text content)
        statusValue = statusCell.textContent.trim() || 'Belum Absen';
      }

      rowData.push(statusValue);

      // Waktu
      const waktuText = cells[5].textContent.trim();
      rowData.push(waktuText === '-' ? '' : waktuText);

      excelData.push(rowData);
    });

    // Validasi apakah ada data yang valid
    if (excelData.length <= 1) {
      alert('Tidak ada data valid untuk diekspor');
      return;
    }

    // Buat workbook menggunakan SheetJS (xlsx)
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(excelData);

    // Set column widths
    const colWidths = [
      { wch: 5 }, // No
      { wch: 15 }, // NIS
      { wch: 25 }, // Nama
      { wch: 10 }, // Kelas
      { wch: 12 }, // Status
      { wch: 20 }, // Waktu
    ];
    ws['!cols'] = colWidths;

    // Set header style (bold)
    const headerRange = XLSX.utils.decode_range(ws['!ref']);
    for (let col = headerRange.s.c; col <= headerRange.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (!ws[cellAddress]) continue;
      ws[cellAddress].s = {
        font: { bold: true },
        fill: { fgColor: { rgb: 'CCCCCC' } },
      };
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Laporan Kehadiran');

    // Generate nama file
    const today = new Date();
    const dateString = today.toISOString().split('T')[0];
    const timeString = today.toTimeString().split(' ')[0].replace(/:/g, '');

    let fileName = 'Laporan_Kehadiran_';
    fileName += kelas === 'all' ? 'Semua_Kelas' : `Kelas_${kelas}`;
    fileName += '_';
    fileName += tanggal.replace(/-/g, '');
    fileName += `_${dateString}_${timeString}.xlsx`;

    // Generate file Excel
    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    // Buat blob dan trigger download
    const blob = new Blob([excelBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(blob);

    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    // Hapus URL objek setelah download
    setTimeout(() => URL.revokeObjectURL(url), 100);

    // Show success message
    if (typeof showAlert === 'function') {
      showAlert(guruAlert, 'success', 'Data berhasil diekspor ke Excel');
    } else {
      alert('Data berhasil diekspor ke Excel');
    }
  } catch (error) {
    console.error('Error saat ekspor:', error);
    alert('Terjadi kesalahan saat mengekspor data: ' + error.message);
  } finally {
    // Reset button state
    exportBtn.disabled = false;
    exportBtn.innerHTML = originalText;
  }
}

// Fungsi alternatif untuk export data langsung dari server (opsional)
function exportToExcelFromServer() {
  const kelas = filterKelas.value;
  const tanggal = filterTanggal.value;
  const exportBtn = document.getElementById('exportBtn');

  if (!tanggal) {
    alert('Pilih tanggal terlebih dahulu');
    return;
  }

  // Format tanggal untuk Apps Script
  const dateParts = tanggal.split('-');
  const formattedDate = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;

  // Tampilkan loading state
  const originalText = exportBtn.innerHTML;
  exportBtn.disabled = true;
  exportBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Mengekspor...';

  const params = new URLSearchParams();
  params.append('action', 'getKehadiranByKelas');
  params.append('kelas', kelas);
  params.append('tanggal', formattedDate);

  fetch(`${SCRIPT_URL}?${params.toString()}`)
    .then((response) => response.json())
    .then((data) => {
      if (data.success && data.data.length > 0) {
        // Siapkan data untuk Excel
        const excelData = [];

        // Header
        excelData.push(['No', 'NIS', 'Nama', 'Kelas', 'Status', 'Waktu']);

        // Data
        data.data.forEach((item, index) => {
          const statusText = item.status || 'Belum Absen';
          const waktuText = item.waktu
            ? new Date(item.waktu).toLocaleDateString('id-ID', {
                weekday: 'short',
                year: 'numeric',
                month: 'short',
                day: 'numeric',
                hour: '2-digit',
                minute: '2-digit',
              })
            : '';

          excelData.push([
            index + 1,
            item.nis,
            item.nama,
            item.kelas,
            statusText,
            waktuText,
          ]);
        });

        // Buat workbook
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(excelData);

        // Set column widths
        const colWidths = [
          { wch: 5 }, // No
          { wch: 15 }, // NIS
          { wch: 25 }, // Nama
          { wch: 10 }, // Kelas
          { wch: 12 }, // Status
          { wch: 20 }, // Waktu
        ];
        ws['!cols'] = colWidths;

        XLSX.utils.book_append_sheet(wb, ws, 'Laporan Kehadiran');

        // Generate nama file
        const today = new Date();
        const dateString = today.toISOString().split('T')[0];

        let fileName = 'Laporan_Kehadiran_';
        fileName += kelas === 'all' ? 'Semua_Kelas' : `Kelas_${kelas}`;
        fileName += '_' + tanggal.replace(/-/g, '');
        fileName += '_' + dateString + '.xlsx';

        // Download file
        const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([excelBuffer], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        const url = URL.createObjectURL(blob);

        const link = document.createElement('a');
        link.href = url;
        link.download = fileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        setTimeout(() => URL.revokeObjectURL(url), 100);

        if (typeof showAlert === 'function') {
          showAlert(guruAlert, 'success', 'Data berhasil diekspor ke Excel');
        } else {
          alert('Data berhasil diekspor ke Excel');
        }
      } else {
        alert('Tidak ada data untuk diekspor');
      }
    })
    .catch((error) => {
      console.error('Error:', error);
      alert('Terjadi kesalahan saat mengekspor data');
    })
    .finally(() => {
      // Reset button state
      exportBtn.disabled = false;
      exportBtn.innerHTML = originalText;
    });
}

function tambahSiswa() {
  const nis = document.getElementById('newNis').value;
  const nama = document.getElementById('newNama').value;
  const kelas = document.getElementById('newKelas').value;
  const password = document.getElementById('newPassword').value;

  if (!nis || !nama || !kelas || !password) {
    showAlert(guruAlert, 'danger', 'Harap isi semua field');
    return;
  }

  const params = new URLSearchParams();
  params.append('action', 'addSiswa');
  params.append('nis', nis);
  params.append('nama', nama);
  params.append('kelas', kelas);
  params.append('password', password);

  fetch(`${SCRIPT_URL}?${params.toString()}`, {
    method: 'POST',
  })
    .then((response) => response.json())
    .then((data) => {
      if (data.success) {
        showAlert(guruAlert, 'success', 'Siswa berhasil ditambahkan');
        // Reset form
        document.getElementById('newNis').value = '';
        document.getElementById('newNama').value = '';
        document.getElementById('newPassword').value = '';
      } else {
        showAlert(
          guruAlert,
          'danger',
          data.message || 'Gagal menambahkan siswa'
        );
      }
    })
    .catch((error) => {
      showAlert(guruAlert, 'danger', 'Terjadi kesalahan. Silakan coba lagi.');
      console.error(error);
    });
}

function tambahKelas() {
  const namaKelas = document.getElementById('namaKelas').value;

  if (!namaKelas) {
    showAlert(guruAlert, 'danger', 'Harap isi nama kelas');
    return;
  }

  const params = new URLSearchParams();
  params.append('action', 'addKelas');
  params.append('namaKelas', namaKelas);

  fetch(`${SCRIPT_URL}?${params.toString()}`, {
    method: 'POST',
  })
    .then((response) => response.json())
    .then((data) => {
      if (data.success) {
        showAlert(guruAlert, 'success', 'Kelas berhasil ditambahkan');
        // Reset form
        document.getElementById('namaKelas').value = '';
        // Reload kelas lists
        loadKelasForFilter();
        loadKelasForAddSiswa();
      } else {
        showAlert(
          guruAlert,
          'danger',
          data.message || 'Gagal menambahkan kelas'
        );
      }
    })
    .catch((error) => {
      showAlert(guruAlert, 'danger', 'Terjadi kesalahan. Silakan coba lagi.');
      console.error(error);
    });
}

function showAlert(element, type, message) {
  element.textContent = message;
  element.classList.remove('hidden', 'alert-success', 'alert-danger');
  element.classList.add(`alert-${type}`);

  // Hide alert after 5 seconds
  setTimeout(() => {
    element.classList.add('hidden');
  }, 5000);
}
