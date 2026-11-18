let dataTerakhir = [];

function formatRupiah(angka, denganTitik = false) {
    const prefix = denganTitik ? "Rp." : "Rp";
    return prefix + angka.toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function hitungAnuitas(P, i, n) {
    const bunga = i / 100;
    let utang = P;
    const data = [];
    let A = 0;

    if (i === 0) {
        A = P / n;
    } else {
        const bBawah = 1 - Math.pow(1 + bunga, -n);
        A = P * (bunga / bBawah);
    }

    for (let periode = 1; periode <= n; periode++) {
        const utangAwal = utang;
        let bungaT = utang * bunga;
        let angsuran = A - bungaT;
        utang = utang - angsuran;

        if (i === 0) {
            bungaT = 0;
            angsuran = A;
        }

        // Indeks: [0]=Periode, [1]=Utang Awal, [2]=Angsuran, [3]=Bunga, [4]=Anuitas, [5]=Sisa Utang
        data.push([periode, utangAwal, angsuran, bungaT, A, utang]);
    }

    return { A, data };
}

function resetClicked() {
    document.getElementById("nilaiP").value = "";
    document.getElementById("nilaiI").value = "";
    document.getElementById("nilaiN").value = "";

    document.getElementById("hasilAnuitas").textContent = "";
    document.getElementById("ringkasanPinjaman").textContent = "";
    document.getElementById("ringkasanBunga").textContent = "";
    document.getElementById("ringkasanTotal").textContent = "";
    document.getElementById("persentaseBunga").textContent = "";
    document.getElementById("pesanError").textContent = "";
    document.getElementById("statusEkspor").textContent = "";
    document.getElementById("tabelBody").innerHTML = "";
    dataTerakhir = [];
}

function hitungClicked() {
    const rawP = document.getElementById("nilaiP").value.trim();
    const rawI = document.getElementById("nilaiI").value.trim();
    const rawN = document.getElementById("nilaiN").value.trim();

    const P = Number(rawP);
    const i = Number(rawI);
    const nFloat = Number(rawN);

    if (rawP === "" || rawI === "" || rawN === "" || isNaN(P) || isNaN(i) || isNaN(nFloat)) {
        resetClicked();
        document.getElementById("pesanError").textContent = "Masukkan input yang valid!";
        return;
    }

    if (P <= 0 || i < 0 || !Number.isInteger(nFloat) || nFloat <= 0) {
        document.getElementById("pesanError").textContent = "Masukkan angka yang benar";
        document.getElementById("hasilAnuitas").textContent = "";
        return;
    }

    document.getElementById("pesanError").textContent = "";
    document.getElementById("statusEkspor").textContent = "";

    const n = Math.trunc(nFloat);
    const { A, data } = hitungAnuitas(P, i, n);
    dataTerakhir = data;

    // row[3] = Bunga, row[4] = Anuitas
    const totalBunga = data.reduce((acc, row) => acc + row[3], 0);
    const totalPembayaran = data.reduce((acc, row) => acc + row[4], 0);

    document.getElementById("ringkasanPinjaman").textContent = formatRupiah(P, false);
    document.getElementById("ringkasanBunga").textContent = formatRupiah(totalBunga, false);
    document.getElementById("ringkasanTotal").textContent = formatRupiah(totalPembayaran, false);
    document.getElementById("persentaseBunga").textContent = `${i}%`;
    document.getElementById("hasilAnuitas").textContent = formatRupiah(A, true);

    const tbody = document.getElementById("tabelBody");
    tbody.innerHTML = data.map(row => `
        <tr>
            <td>${row[0]}</td>
            <td>${formatRupiah(row[1], true)}</td>
            <td>${formatRupiah(row[2], true)}</td>
            <td>${formatRupiah(row[3], true)}</td>
            <td>${formatRupiah(row[4], true)}</td>
            <td>${formatRupiah(Math.abs(row[5]) < 1e-9 ? 0 : row[5], true)}</td>
        </tr>
    `).join("");
}

function eksporClicked() {
    const statusEl = document.getElementById("statusEkspor");
    statusEl.style.color = "#ef4444";

    if (dataTerakhir.length === 0) {
        statusEl.textContent = "Hitung anuitas terlebih dahulu!";
        return;
    }

    const header = [
        "Periode",
        "Pinjaman Awal",
        "Angsuran Pokok",
        "Bunga",
        "Anuitas",
        "Sisa Utang"
    ];

    const wsData = [header, ...dataTerakhir];
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Format mata uang Rupiah di Excel
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = 1; R <= range.e.r; ++R) {
        for (let C = 1; C <= range.e.c; ++C) {
            const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
            if (ws[cellRef]) {
                ws[cellRef].z = '"Rp"#,##0.00';
            }
        }
    }

    ws['!cols'] = [
        { wch: 10 },
        { wch: 20 },
        { wch: 20 },
        { wch: 18 },
        { wch: 20 },
        { wch: 20 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rincian Anuitas");
    XLSX.writeFile(wb, "Rincian_Anuitas.xlsx");

    statusEl.style.color = "#16a34a";
    statusEl.textContent = "Proses ekspor selesai.";
}
