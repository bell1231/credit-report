const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const DATA_DIR = path.join(__dirname, '.data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

const USERS_FILE = path.join(DATA_DIR, 'users.json');
const SECRET_FILE = path.join(DATA_DIR, '.secret');

let TOKEN_SECRET;
if (fs.existsSync(SECRET_FILE)) {
  TOKEN_SECRET = fs.readFileSync(SECRET_FILE, 'utf-8').trim();
} else {
  TOKEN_SECRET = crypto.randomBytes(32).toString('hex');
  fs.writeFileSync(SECRET_FILE, TOKEN_SECRET);
}

function hashPassword(password) {
  return crypto.createHash('sha256').update(password + TOKEN_SECRET).digest('hex');
}

const users = {
  "admin@credit-report.com": {
    email: "admin@credit-report.com",
    passwordHash: hashPassword("Admin@123456"),
    name: "管理员",
    createdAt: Date.now(),
  },
  "demo@credit-report.com": {
    email: "demo@credit-report.com",
    passwordHash: hashPassword("Demo@123456"),
    name: "演示用户",
    createdAt: Date.now(),
  },
};

fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2));
console.log('用户数据已初始化:', USERS_FILE);
console.log('测试账户:');
console.log('  admin@credit-report.com / Admin@123456');
console.log('  demo@credit-report.com  / Demo@123456');
