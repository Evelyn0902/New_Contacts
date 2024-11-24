const express = require('express');
const mysql = require('mysql');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');

const app = express();
const port = 5000;

// 使用环境变量来管理数据库配置
const dbConfig = {
  host: process.env.DB_HOST || 'localhost',
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || '123456',
  database: process.env.DB_NAME || 'contacts'
};

const connection = mysql.createConnection(dbConfig);

connection.connect(err => {
  if (err) {
    console.error('数据库连接失败: ' + err.stack);
    return;
  }
  console.log('数据库连接成功，ID为: ' + connection.threadId);
});

app.use(express.static('pages'));
app.use(bodyParser.json());

// 路由处理
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/pages/index.html');
});

app.post('/add', async (req, res) => {
  const { name, phone, email, address } = req.body;
  if (!name || !phone || !email || !address) {
    return res.status(400).send('所有字段都是必填项');
  }

  const query = 'INSERT INTO contacts (name, phone, email, address) VALUES (?, ?, ?, ?)';
  connection.query(query, [name, phone, email, address], (error, results) => {
    if (error) {
      console.error('添加数据时出错:', error.message);
      return res.status(500).send('内部服务器错误');
    }
    res.send('ok');
  });
});

app.get('/contacts', (req, res) => {
  connection.query('SELECT * FROM contacts', (error, results) => {
    if (error) {
      console.error('获取数据时出错:', error.message);
      return res.status(500).send('内部服务器错误');
    }
    res.json(results);
  });
});

app.post('/del', (req, res) => {
  const { id } = req.body;
  if (!id) {
    return res.status(400).send('ID是必填项');
  }

  const deleteSql = 'DELETE FROM contacts WHERE id = ?';
  connection.query(deleteSql, [id], (error, results) => {
    if (error) {
      console.error('删除数据时出错:', error.message);
      return res.status(500).send('内部服务器错误');
    }
    res.send('ok');
  });
});

app.post('/edit', (req, res) => {
  const { id, name, phone, newphone, email, address } = req.body;
  if (!id || !name || !email || !address) {
    return res.status(400).send('所有字段都是必填项');
  }

  connection.query('SELECT * FROM contacts WHERE id = ?', [id], (error, results) => {
    if (error) {
      console.error('更新数据时出错:', error.message);
      return res.status(500).send('内部服务器错误');
    }

    let newp = phone;
    if (newphone) {
      const currentPhones = results[0].phone ? results[0].phone.split(';') : [];
      if (!currentPhones.includes(phone)) {
        newp = `${phone};${results[0].phone}`;
      }
    }

    const updateQuery = 'UPDATE contacts SET name = ?, phone = ?, email = ?, address = ? WHERE id = ?';
    connection.query(updateQuery, [name, newp, email, address, id], (error, results) => {
      if (error) {
        console.error('更新数据时出错:', error.message);
        return res.status(500).send('内部服务器错误');
      }
      res.send('ok');
    });
  });
});

app.post('/output', (req, res) => {
  const data = req.body;
  const workbook = XLSX.utils.book_new();
  const xlsdata = [['id', '姓名', '电话', '邮箱', '地址']];
  data.forEach(item => {
    xlsdata.push([item.id, item.name, item.phone, item.email, item.address]);
  });
  const worksheet = XLSX.utils.aoa_to_sheet(xlsdata);
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
  XLSX.writeFile(workbook, "./pages/output.xlsx");
  res.send('ok');
});

app.post('/input', (req, res) => {
  const data = req.body;
  data.forEach(item => {
    const query = 'INSERT INTO contacts (name, phone, email, address) VALUES (?, ?, ?, ?)';
    connection.query(query, [item['姓名'], item['电话'], item['邮箱'], item['地址']], (error, results) => {
      if (error) {
        console.error('导入数据时出错:', error.message);
        return res.status(500).send('内部服务器错误');
      }
    });
  });
  res.send('ok');
});

// 错误处理中间件
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).send('内部服务器错误');
});

app.listen(port, () => {
  console.log(`服务在http://127.0.0.1:${port}/ 启动`);
});