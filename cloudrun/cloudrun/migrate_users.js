// 迁移用户数据到 CloudBase 数据库
const cloudbase = require('@cloudbase/node-sdk');
const fs = require('fs');
const path = require('path');

const app = cloudbase.init({ env: 'sihui2026-d6gtqfgdw7803a003' });
const db = app.database();

const USERS_FILE = path.join(__dirname, '.data', 'users.json');

async function migrate() {
  // 读取本地用户数据
  let users = {};
  if (fs.existsSync(USERS_FILE)) {
    users = JSON.parse(fs.readFileSync(USERS_FILE, 'utf-8'));
  }
  console.log(`本地用户数: ${Object.keys(users).length}`);

  let inserted = 0, updated = 0;

  for (const [email, userData] of Object.entries(users)) {
    try {
      // 检查是否已存在
      const existing = await db.collection('users').where({ email }).limit(1).get();
      if (existing.data && existing.data.length > 0) {
        // 更新
        await db.collection('users').doc(existing.data[0]._id).update({
          passwordHash: userData.passwordHash,
          name: userData.name || email.split('@')[0],
          updatedAt: Date.now(),
        });
        updated++;
        console.log(`  更新: ${email}`);
      } else {
        // 插入
        await db.collection('users').add({
          email,
          passwordHash: userData.passwordHash,
          name: userData.name || email.split('@')[0],
          createdAt: userData.createdAt || Date.now(),
          updatedAt: Date.now(),
        });
        inserted++;
        console.log(`  新增: ${email}`);
      }
    } catch (e) {
      console.error(`  失败: ${email} - ${e.message}`);
    }
  }

  console.log(`\n迁移完成: 新增 ${inserted}, 更新 ${updated}`);
}

migrate().catch(console.error);
