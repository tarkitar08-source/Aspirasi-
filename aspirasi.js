document.addEventListener("DOMContentLoaded", () => {
    // KODE RAHASIA UNTUK DOWNLOAD EXCEL (Dapat diubah sesuai keinginan pengurus)
    const CORRECT_CODE = "osamasa2026"; 

    // 1. Logika Pengiriman Aspirasi Kelas
    const form = document.getElementById("classmeet-form");
    const inputKelas = document.getElementById("input-kelas");
    const inputAspirasi = document.getElementById("input-aspirasi");
    const successAlert = document.getElementById("success-alert");

    if (form) {
        form.addEventListener("submit", (e) => {
            e.preventDefault();

            const kelasVal = inputKelas.value.trim();
            const aspirasiVal = inputAspirasi.value.trim();
            const waktuVal = new Date().toLocaleString("id-ID", { 
                day: '2-digit', month: 'short', year: 'numeric', 
                hour: '2-digit', minute: '2-digit' 
            });

            const newEntry = {
                no: 0, // Akan diurutkan otomatis saat export
                kelas: kelasVal,
                aspirasi: aspirasiVal,
                waktu: waktuVal
            };

            // Ambil data lama dari localStorage atau buat array baru
            let existingData = JSON.parse(localStorage.getItem("classmeet_aspirasi_db")) || [];
            existingData.push(newEntry);
            localStorage.setItem("classmeet_aspirasi_db", JSON.stringify(existingData));

            // Reset Form & Tampilkan Notifikasi
            form.reset();
            successAlert.classList.remove("hidden");
            setTimeout(() => {
                successAlert.classList.add("hidden");
            }, 4000);
        });
    }

    // 2. Logika Modal Password & Download Excel
    const openModalBtn = document.getElementById("open-auth-modal");
    const closeModalBtn = document.getElementById("close-modal");
    const authModal = document.getElementById("auth-modal");
    const authForm = document.getElementById("auth-form");
    const secretCodeInput = document.getElementById("secret-code");
    const authError = document.getElementById("auth-error");

    if (openModalBtn) {
        openModalBtn.addEventListener("click", () => {
            authModal.classList.remove("hidden");
            secretCodeInput.value = "";
            authError.classList.add("hidden");
        });
    }

    if (closeModalBtn) {
        closeModalBtn.addEventListener("click", () => {
            authModal.classList.add("hidden");
        });
    }

    if (authForm) {
        authForm.addEventListener("submit", (e) => {
            e.preventDefault();
            const enteredCode = secretCodeInput.value.trim();

            if (enteredCode === CORRECT_CODE) {
                authModal.classList.add("hidden");
                downloadExcelFile();
            } else {
                authError.classList.remove("hidden");
            }
        });
    }

    // 3. Fungsi Membuat dan Mengunduh File Excel yang Rapih menggunakan SheetJS
    function downloadExcelFile() {
        let storedData = JSON.parse(localStorage.getItem("classmeet_aspirasi_db")) || [];

        if (storedData.length === 0) {
            alert("Belum ada data aspirasi kelas yang masuk.");
            return;
        }

        // Format data agar tabel rapi sesuai permintaan (Kolom: No, Nama Kelas, Aspirasi, Waktu)
        let formattedData = storedData.map((item, index) => ({
            "No": index + 1,
            "Nama Kelas": item.kelas,
            "Aspirasi / Usulan Kegiatan Classmeet": item.aspirasi,
            "Waktu Pengiriman": item.waktu
        }));

        // Buat Worksheet & Workbook
        const worksheet = XLSX.utils.json_to_sheet(formattedData);
        
        // Atur lebar kolom agar rapi dan tidak terpotong
        worksheet['!cols'] = [
            { wch: 6 },  // No
            { wch: 15 }, // Nama Kelas
            { wch: 55 }, // Aspirasi
            { wch: 20 }  // Waktu
        ];

        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "Aspirasi Classmeet");

        // Unduh File Excel (.xlsx) ke perangkat
        XLSX.writeFile(workbook, "Rekap_Aspirasi_Classmeet_OSAMASA.xlsx");
    }
});