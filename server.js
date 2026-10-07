import express from 'express';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import multer from 'multer';
import methodOverride from 'method-override';
import session from 'express-session';
import flash from 'connect-flash';
import ExcelJS from 'exceljs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = '0.0.0.0';

// Ensure upload directory exists
const uploadDir = path.join(__dirname, 'uploads', 'lampiran');
fs.mkdirSync(uploadDir, { recursive: true });

// Create sample lampiran files if not existing
const sampleFiles = [
    { name: 'sample_sj_001.pdf', content: 'Sample Surat Jalan SJ-2026-001 Document Content' },
    { name: 'sample_sj_002.pdf', content: 'Sample Surat Jalan SJ-2026-002 Document Content' },
    { name: 'sample_sj_003.png', content: 'Sample Image Content' },
];
for (const sf of sampleFiles) {
    const p = path.join(uploadDir, sf.name);
    if (!fs.existsSync(p)) {
        fs.writeFileSync(p, sf.content);
    }
}

// Multer setup for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        const safe = Date.now() + '_' + file.originalname.replace(/[^a-zA-Z0-9._-]/g, '_');
        cb(null, safe);
    }
});
const upload = multer({
    storage,
    limits: { fileSize: 10 * 1024 * 1024 } // 10MB
});

// App configuration
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(methodOverride('_method'));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(path.join(__dirname, 'uploads')));

// Session and Flash middleware
app.use(session({
    secret: process.env.SESSION_SECRET || 'kurang-kirim-secret-key-2026',
    resave: false,
    saveUninitialized: false,
    cookie: { maxAge: 60000 * 60 }
}));
app.use(flash());

// Helper for formatting date dd/mm/yyyy
function formatDate(d) {
    if (!d) return '-';
    if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(d)) {
        const [y, m, day] = d.split('-');
        return `${day}/${m}/${y}`;
    }
    const dt = new Date(d);
    if (isNaN(dt.getTime())) return String(d);
    const day = String(dt.getDate()).padStart(2, '0');
    const month = String(dt.getMonth() + 1).padStart(2, '0');
    const year = dt.getFullYear();
    return `${day}/${month}/${year}`;
}

// Global view variables
app.use((req, res, next) => {
    res.locals.currentPath = req.path;
    res.locals.success = req.flash('success')[0] || null;
    res.locals.error = req.flash('error')[0] || null;
    res.locals.formatDate = formatDate;
    res.locals.stats = {
        totalSj: kurangKirimList ? kurangKirimList.filter(k => k.status === 1).length : 0,
        totalToko: tokoList ? tokoList.length : 0,
        activeToko: tokoList ? tokoList.filter(t => t.status === 1).length : 0
    };
    next();
});

// -------------------------------------------------------------
// In-Memory Database (Pre-seeded with Toko and Sample Records)
// -------------------------------------------------------------
let tokoNextId = 11;
const tokoList = [
    { id: 1, kode_toko: 'TK001', nama_toko: 'Toko Sejahtera', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 2, kode_toko: 'TK002', nama_toko: 'Toko Makmur', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 3, kode_toko: 'TK003', nama_toko: 'Toko Maju Jaya', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 4, kode_toko: 'TK004', nama_toko: 'Toko Berkah', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 5, kode_toko: 'TK005', nama_toko: 'Toko Sentosa', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 6, kode_toko: 'TK006', nama_toko: 'Toko Harapan', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 7, kode_toko: 'TK007', nama_toko: 'Toko Abadi', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 8, kode_toko: 'TK008', nama_toko: 'Toko Mandiri', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 9, kode_toko: 'TK009', nama_toko: 'Toko Bersama', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
    { id: 10, kode_toko: 'TK010', nama_toko: 'Toko Gemilang', status: 1, created_at: new Date('2026-04-20T08:00:00Z') },
];

let kurangKirimNextId = 4;
const kurangKirimList = [
    {
        id: 1,
        toko_id: 1,
        kode_toko: 'TK001',
        tgl_kirim: '2026-05-01',
        nomor_surat_jalan: 'SJ-2026-001',
        lampiran: 'sample_sj_001.pdf',
        original_filename: 'surat_jalan_TK001.pdf',
        status: 1,
        created_at: new Date('2026-05-01T09:30:00Z')
    },
    {
        id: 2,
        toko_id: 3,
        kode_toko: 'TK003',
        tgl_kirim: '2026-05-03',
        nomor_surat_jalan: 'SJ-2026-002',
        lampiran: 'sample_sj_002.pdf',
        original_filename: 'surat_jalan_TK003.pdf',
        status: 1,
        created_at: new Date('2026-05-03T14:20:00Z')
    },
    {
        id: 3,
        toko_id: 5,
        kode_toko: 'TK005',
        tgl_kirim: '2026-05-04',
        nomor_surat_jalan: 'SJ-2026-003',
        lampiran: 'sample_sj_003.png',
        original_filename: 'bukti_kurang_TK005.png',
        status: 1,
        created_at: new Date('2026-05-04T11:45:00Z')
    }
];

const uploadLogs = [];

// Helper to get active stores sorted by kode_toko
function getActiveTokoList() {
    return tokoList
        .filter(t => t.status === 1)
        .sort((a, b) => a.kode_toko.localeCompare(b.kode_toko));
}

// -------------------------------------------------------------
// Routes
// -------------------------------------------------------------

// Root redirect
app.get('/', (req, res) => {
    res.redirect('/dashboard');
});

// Dashboard: Input Surat Jalan Form
app.get('/dashboard', (req, res) => {
    const activeToko = getActiveTokoList();
    const today = new Date().toISOString().split('T')[0];
    res.render('dashboard', {
        tokoList: activeToko,
        defaultDate: today
    });
});

// Store Surat Jalan
app.post('/dashboard', upload.single('lampiran'), (req, res) => {
    try {
        const { toko_id, tgl_kirim, nomor_surat_jalan } = req.body;

        if (!toko_id || !tgl_kirim || !nomor_surat_jalan) {
            req.flash('error', 'Semua kolom bertanda * wajib diisi.');
            return res.redirect('/dashboard');
        }

        const toko = tokoList.find(t => t.id === parseInt(toko_id, 10));
        if (!toko) {
            req.flash('error', 'Toko yang dipilih tidak valid.');
            return res.redirect('/dashboard');
        }

        let savedFilename = null;
        let originalFilename = null;
        if (req.file) {
            savedFilename = req.file.filename;
            originalFilename = req.file.originalname;

            uploadLogs.push({
                id: uploadLogs.length + 1,
                original_filename: req.file.originalname,
                safe_filename: req.file.filename,
                file_size: req.file.size,
                mime_type: req.file.mimetype,
                upload_path: path.join('lampiran', req.file.filename),
                ip_address: req.ip,
                created_at: new Date()
            });
        }

        const newRecord = {
            id: kurangKirimNextId++,
            toko_id: toko.id,
            kode_toko: toko.kode_toko,
            tgl_kirim,
            nomor_surat_jalan,
            lampiran: savedFilename,
            original_filename: originalFilename,
            status: 1,
            created_at: new Date()
        };

        kurangKirimList.unshift(newRecord);

        req.flash('success', 'Surat Jalan berhasil disimpan!');
        res.redirect('/dashboard');
    } catch (err) {
        console.error('Error saving Surat Jalan:', err);
        req.flash('error', 'Gagal menyimpan Surat Jalan: ' + err.message);
        res.redirect('/dashboard');
    }
});

// Helper for filtering & sorting Riwayat
function getFilteredKurangKirim(query) {
    let list = kurangKirimList
        .filter(item => item.status === 1)
        .map(item => {
            const toko = tokoList.find(t => t.id === item.toko_id);
            return { ...item, toko };
        });

    // Search filter
    if (query.search && query.search.trim()) {
        const s = query.search.trim().toLowerCase();
        list = list.filter(item => {
            const sj = (item.nomor_surat_jalan || '').toLowerCase();
            const kt = (item.kode_toko || '').toLowerCase();
            const nt = (item.toko && item.toko.nama_toko ? item.toko.nama_toko : '').toLowerCase();
            return sj.includes(s) || kt.includes(s) || nt.includes(s);
        });
    }

    // Date filters
    if (query.date_from) {
        list = list.filter(item => item.tgl_kirim >= query.date_from);
    }
    if (query.date_to) {
        list = list.filter(item => item.tgl_kirim <= query.date_to);
    }

    // Sorting
    const sortField = query.sort || 'created_at';
    const sortDirection = query.direction === 'asc' ? 'asc' : 'desc';

    list.sort((a, b) => {
        let valA = a[sortField];
        let valB = b[sortField];

        if (sortField === 'created_at') {
            valA = new Date(valA).getTime();
            valB = new Date(valB).getTime();
        } else if (typeof valA === 'string') {
            valA = valA.toLowerCase();
            valB = (valB || '').toLowerCase();
        }

        if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
        if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
        return 0;
    });

    return { list, sortField, sortDirection };
}

// History / Riwayat
app.get('/history', (req, res) => {
    const { list, sortField, sortDirection } = getFilteredKurangKirim(req.query);

    const currentPage = parseInt(req.query.page, 10) || 1;
    const perPage = 10;
    const totalItems = list.length;
    const totalPages = Math.ceil(totalItems / perPage) || 1;
    const paginatedData = list.slice((currentPage - 1) * perPage, currentPage * perPage);

    // Helpers for building URLs
    function buildSortUrl(field, direction) {
        const params = new URLSearchParams(req.query);
        params.set('sort', field);
        params.set('direction', direction);
        return params.toString();
    }

    function buildPageUrl(page) {
        const params = new URLSearchParams(req.query);
        params.set('page', page);
        return params.toString();
    }

    res.render('history', {
        dataList: paginatedData,
        totalItems,
        currentPage,
        totalPages,
        perPage,
        sortField,
        sortDirection,
        query: req.query,
        queryString: new URLSearchParams(req.query).toString(),
        buildSortUrl,
        buildPageUrl
    });
});

// Export Excel
app.get('/history/export', async (req, res) => {
    try {
        const { list } = getFilteredKurangKirim(req.query);

        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Kurang Kirim';
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet('Kurang Kirim');

        // Define columns
        worksheet.columns = [
            { header: 'No', key: 'no', width: 8 },
            { header: 'Kode Toko', key: 'kode_toko', width: 16 },
            { header: 'Nama Toko', key: 'nama_toko', width: 30 },
            { header: 'Tanggal Kirim', key: 'tgl_kirim', width: 18 },
            { header: 'Nomor Surat Jalan', key: 'nomor_surat_jalan', width: 25 },
            { header: 'File Lampiran', key: 'lampiran', width: 30 },
        ];

        // Style Header row (purple #6D28D9, white bold text, centered)
        const headerRow = worksheet.getRow(1);
        headerRow.height = 25;
        headerRow.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF6D28D9' }
            };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
        });

        // Add Data Rows
        list.forEach((item, index) => {
            worksheet.addRow({
                no: index + 1,
                kode_toko: item.kode_toko,
                nama_toko: item.toko ? item.toko.nama_toko : '-',
                tgl_kirim: formatDate(item.tgl_kirim),
                nomor_surat_jalan: item.nomor_surat_jalan,
                lampiran: item.original_filename || item.lampiran || '-'
            });
        });

        const filename = `kurang_kirim_${new Date().toISOString().slice(0, 10)}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);

        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error('Error exporting Excel:', err);
        req.flash('error', 'Gagal mengekspor data: ' + err.message);
        res.redirect('/history');
    }
});

// Edit Surat Jalan Form
app.get('/kurang-kirim/:id/edit', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const item = kurangKirimList.find(k => k.id === id && k.status === 1);

    if (!item) {
        req.flash('error', 'Data tidak ditemukan.');
        return res.redirect('/history');
    }

    const activeToko = getActiveTokoList();
    res.render('edit', {
        item,
        tokoList: activeToko
    });
});

// Update Surat Jalan
app.put('/kurang-kirim/:id', upload.single('lampiran'), (req, res) => {
    const id = parseInt(req.params.id, 10);
    const item = kurangKirimList.find(k => k.id === id && k.status === 1);

    if (!item) {
        req.flash('error', 'Data tidak ditemukan.');
        return res.redirect('/history');
    }

    const { toko_id, tgl_kirim, nomor_surat_jalan } = req.body;
    const toko = tokoList.find(t => t.id === parseInt(toko_id, 10));

    if (!toko) {
        req.flash('error', 'Toko tidak valid.');
        return res.redirect(`/kurang-kirim/${id}/edit`);
    }

    item.toko_id = toko.id;
    item.kode_toko = toko.kode_toko;
    item.tgl_kirim = tgl_kirim;
    item.nomor_surat_jalan = nomor_surat_jalan;

    if (req.file) {
        item.lampiran = req.file.filename;
        item.original_filename = req.file.originalname;

        uploadLogs.push({
            id: uploadLogs.length + 1,
            original_filename: req.file.originalname,
            safe_filename: req.file.filename,
            file_size: req.file.size,
            mime_type: req.file.mimetype,
            upload_path: path.join('lampiran', req.file.filename),
            ip_address: req.ip,
            created_at: new Date()
        });
    }

    req.flash('success', 'Data berhasil diperbarui!');
    res.redirect('/history');
});

// Soft-delete Surat Jalan
app.delete('/kurang-kirim/:id', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const item = kurangKirimList.find(k => k.id === id);

    if (item) {
        item.status = 0;
        req.flash('success', 'Data berhasil dihapus!');
    } else {
        req.flash('error', 'Data tidak ditemukan.');
    }

    res.redirect('/history');
});

// View Lampiran inline
app.get('/kurang-kirim/:id/view', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const item = kurangKirimList.find(k => k.id === id);

    if (!item || !item.lampiran) {
        return res.status(404).send('Lampiran tidak ditemukan.');
    }

    const filePath = path.join(uploadDir, path.basename(item.lampiran));
    if (fs.existsSync(filePath)) {
        return res.sendFile(filePath);
    }

    // Fallback placeholder response
    res.setHeader('Content-Type', 'text/plain');
    res.send(`Lampiran file: ${item.original_filename || item.lampiran}\nNomor Surat Jalan: ${item.nomor_surat_jalan}`);
});

// Download Lampiran
app.get('/kurang-kirim/:id/download', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const item = kurangKirimList.find(k => k.id === id);

    if (!item || !item.lampiran) {
        return res.status(404).send('Lampiran tidak ditemukan.');
    }

    const filePath = path.join(uploadDir, path.basename(item.lampiran));
    const downloadName = item.original_filename || path.basename(item.lampiran);

    if (fs.existsSync(filePath)) {
        return res.download(filePath, downloadName);
    }

    res.setHeader('Content-Disposition', `attachment; filename="${downloadName}"`);
    res.setHeader('Content-Type', 'text/plain');
    res.send(`Surat Jalan: ${item.nomor_surat_jalan}\nKode Toko: ${item.kode_toko}\nTanggal: ${item.tgl_kirim}`);
});

// -------------------------------------------------------------
// Master Toko Routes
// -------------------------------------------------------------

// Toko List
app.get('/toko', (req, res) => {
    let list = [...tokoList];
    const search = req.query.search ? req.query.search.trim().toLowerCase() : '';

    if (search) {
        list = list.filter(t => 
            t.kode_toko.toLowerCase().includes(search) || 
            t.nama_toko.toLowerCase().includes(search)
        );
    }

    list.sort((a, b) => a.kode_toko.localeCompare(b.kode_toko));

    const currentPage = parseInt(req.query.page, 10) || 1;
    const perPage = 20;
    const totalItems = list.length;
    const totalPages = Math.ceil(totalItems / perPage) || 1;
    const paginatedData = list.slice((currentPage - 1) * perPage, currentPage * perPage);

    res.render('toko/index', {
        tokoList: paginatedData,
        search: req.query.search || '',
        totalItems,
        currentPage,
        totalPages,
        perPage
    });
});

// Toko Create Form
app.get('/toko/create', (req, res) => {
    res.render('toko/create');
});

// Toko Store
app.post('/toko', (req, res) => {
    const { kode_toko, nama_toko } = req.body;

    if (!kode_toko || !nama_toko) {
        req.flash('error', 'Kode Toko dan Nama Toko wajib diisi.');
        return res.redirect('/toko/create');
    }

    const existing = tokoList.find(t => t.kode_toko.toLowerCase() === kode_toko.trim().toLowerCase());
    if (existing) {
        req.flash('error', `Kode Toko ${kode_toko} sudah digunakan.`);
        return res.redirect('/toko/create');
    }

    const newToko = {
        id: tokoNextId++,
        kode_toko: kode_toko.trim().toUpperCase(),
        nama_toko: nama_toko.trim(),
        status: 1,
        created_at: new Date()
    };
    tokoList.push(newToko);

    req.flash('success', 'Toko berhasil ditambahkan!');
    res.redirect('/toko');
});

// Toko Edit Form
app.get('/toko/:id/edit', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const toko = tokoList.find(t => t.id === id);

    if (!toko) {
        req.flash('error', 'Toko tidak ditemukan.');
        return res.redirect('/toko');
    }

    res.render('toko/edit', { toko });
});

// Toko Update
app.put('/toko/:id', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const toko = tokoList.find(t => t.id === id);

    if (!toko) {
        req.flash('error', 'Toko tidak ditemukan.');
        return res.redirect('/toko');
    }

    const { kode_toko, nama_toko } = req.body;
    if (!kode_toko || !nama_toko) {
        req.flash('error', 'Kode Toko dan Nama Toko wajib diisi.');
        return res.redirect(`/toko/${id}/edit`);
    }

    const duplicate = tokoList.find(t => t.id !== id && t.kode_toko.toLowerCase() === kode_toko.trim().toLowerCase());
    if (duplicate) {
        req.flash('error', `Kode Toko ${kode_toko} sudah digunakan oleh toko lain.`);
        return res.redirect(`/toko/${id}/edit`);
    }

    toko.kode_toko = kode_toko.trim().toUpperCase();
    toko.nama_toko = nama_toko.trim();

    req.flash('success', 'Toko berhasil diperbarui!');
    res.redirect('/toko');
});

// Toko Delete
app.delete('/toko/:id', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const toko = tokoList.find(t => t.id === id);

    if (!toko) {
        req.flash('error', 'Toko tidak ditemukan.');
        return res.redirect('/toko');
    }

    // Check if active surat jalan exist for this toko
    const activeSuratJalanCount = kurangKirimList.filter(k => k.toko_id === id && k.status === 1).length;
    if (activeSuratJalanCount > 0) {
        req.flash('error', `Toko tidak bisa dihapus karena masih ada ${activeSuratJalanCount} surat jalan yang terikat.`);
        return res.redirect('/toko');
    }

    const index = tokoList.findIndex(t => t.id === id);
    if (index !== -1) {
        tokoList.splice(index, 1);
        req.flash('success', 'Toko berhasil dihapus!');
    }

    res.redirect('/toko');
});

// Toggle Toko Status
app.patch('/toko/:id/toggle-status', (req, res) => {
    const id = parseInt(req.params.id, 10);
    const toko = tokoList.find(t => t.id === id);

    if (!toko) {
        req.flash('error', 'Toko tidak ditemukan.');
        return res.redirect('/toko');
    }

    toko.status = toko.status ? 0 : 1;
    const statusLabel = toko.status ? 'diaktifkan' : 'dinonaktifkan';

    req.flash('success', `Toko ${toko.nama_toko} berhasil ${statusLabel}!`);
    res.redirect('/toko');
});

// Start Server
app.listen(PORT, HOST, () => {
    console.log(`Kurang Kirim application running at http://${HOST}:${PORT}`);
});
