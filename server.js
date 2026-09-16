const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');

const PORT = 3000;
const DB_FILE = path.join(__dirname, 'database.json');

// التأكد من وجود ملف قاعدة البيانات
if (!fs.existsSync(DB_FILE)) {
  fs.writeFileSync(DB_FILE, JSON.stringify([]));
}

// دالة لجلب الآي بي الخاص بالجهاز في الشبكة المحلية
function getLocalIp() {
  const interfaces = os.networkInterfaces();
  for (const name in interfaces) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) {
        return iface.address;
      }
    }
  }
  return 'localhost';
}

const server = http.createServer((req, res) => {
  // تمكين CORS للسماح بالاتصال من صفحات الويب المفتوحة كملفات محلية file://
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const urlPath = req.url.split('?')[0];

  // التحقق هل الطلب هو لجلب البيانات أو حفظها (API)
  const accept = req.headers['accept'] || '';
  const isFetch = req.headers['sec-fetch-dest'] === 'empty' || accept.includes('application/json') || !accept.includes('text/html');

  if (urlPath === '/' && isFetch) {
    // ----------------------------------------------------
    // GET API - جلب جميع السجلات
    // ----------------------------------------------------
    if (req.method === 'GET') {
      fs.readFile(DB_FILE, 'utf8', (err, data) => {
        if (err) {
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ success: false, error: 'Failed to read database' }));
          return;
        }
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(data);
      });
      return;
    }
    // ----------------------------------------------------
    // POST API - حفظ/تحديث/حذف السجلات
    // ----------------------------------------------------
    else if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => { body += chunk; });
      req.on('end', () => {
        try {
          const postData = JSON.parse(body);
          
          fs.readFile(DB_FILE, 'utf8', (err, dbContent) => {
            if (err) throw err;
            let records = JSON.parse(dbContent || '[]');

            // حالة الحذف
            if (postData.action === 'delete') {
              const idToDelete = postData.id;
              const originalLength = records.length;
              records = records.filter(r => r.id.toString() !== idToDelete.toString());
              
              if (records.length === originalLength) {
                res.writeHead(404, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ success: false, message: 'السجل غير موجود' }));
                return;
              }

              fs.writeFile(DB_FILE, JSON.stringify(records, null, 2), 'utf8', (writeErr) => {
                if (writeErr) throw writeErr;
                res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
                res.end(JSON.stringify({ success: true, message: 'تم الحذف بنجاح' }));
              });
              return;
            }

            // حالة الحفظ أو التعديل
            const record = postData;
            const recordId = record.id;
            const foundIndex = records.findIndex(r => r.id.toString() === recordId.toString());

            if (foundIndex !== -1) {
              // تحديث
              records[foundIndex] = record;
            } else {
              // إضافة جديد
              records.push(record);
            }

            fs.writeFile(DB_FILE, JSON.stringify(records, null, 2), 'utf8', (writeErr) => {
              if (writeErr) throw writeErr;
              res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
              res.end(JSON.stringify({ success: true, message: 'تم الحفظ والتحديث بنجاح' }));
            });
          });
        } catch (parseErr) {
          res.writeHead(400, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ success: false, error: 'Invalid JSON data' }));
        }
      });
      return;
    }
  }

  // ----------------------------------------------------
  // خدمة الملفات الثابتة (Static Files)
  // ----------------------------------------------------
  if (req.method === 'GET') {
    let filePath = path.join(__dirname, urlPath === '/' ? 'index.html' : urlPath);
    
    // منع الخروج من المجلد لضمان الحماية
    if (!filePath.startsWith(__dirname)) {
      res.writeHead(403);
      res.end('Forbidden');
      return;
    }

    const extname = String(path.extname(filePath)).toLowerCase();
    const mimeTypes = {
      '.html': 'text/html; charset=utf-8',
      '.js': 'text/javascript; charset=utf-8',
      '.css': 'text/css; charset=utf-8',
      '.json': 'application/json; charset=utf-8',
      '.png': 'image/png',
      '.jpg': 'image/jpg',
      '.gif': 'image/gif',
      '.svg': 'image/svg+xml'
    };

    const contentType = mimeTypes[extname] || 'application/octet-stream';

    fs.readFile(filePath, (err, content) => {
      if (err) {
        if (err.code === 'ENOENT') {
          res.writeHead(404, { 'Content-Type': 'text/html; charset=utf-8' });
          res.end('<h1>404 - الملف غير موجود</h1>', 'utf-8');
        } else {
          res.writeHead(500);
          res.end(`Server Error: ${err.code}`);
        }
      } else {
        res.writeHead(200, { 'Content-Type': contentType });
        res.end(content, 'utf-8');
      }
    });
  } else {
    res.writeHead(405);
    res.end('Method Not Allowed');
  }
});

server.listen(PORT, () => {
  const localIp = getLocalIp();
  console.log('==================================================');
  console.log('    نظام إدارة النشاطات - تشغيل الخادم المحلي بنجاح');
  console.log('==================================================');
  console.log(`- للوصول من حاسوبك الشخصي: http://localhost:${PORT}`);
  console.log(`- للوصول من أجهزة الموظفين في الشبكة المحلية: http://${localIp}:${PORT}`);
  console.log('==================================================');
  console.log('اضغط Ctrl + C لإيقاف الخادم.');
});
