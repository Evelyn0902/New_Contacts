const express = require('express')

const mysql = require('mysql')

const bodyparser = require('body-parser')

const XLSX = require('xlsx');

const app = express()

const connection = mysql.createConnection({
    host: 'localhost', // 数据库服务器地址
    user: 'root', // 数据库用户名
    password: '123456', // 数据库密码
    database: 'contacts' // 要连接的数据库名
})

connection.connect()

app.use(express.static('pages'))
app.use(bodyparser.json())

app.get('/', (req, res) => {
    // 假设你的HTML文件名为index.html，位于public目录下
    res.sendFile(__dirname + '/pages/index.html')
})

app.post('/add', (req, res) => {
    let info = req.body
    console.log(info)
    const query =
        'INSERT INTO contacts (name, phone, email, address) VALUES (?, ?, ?, ?)'
    let values = [info.name, info.phone, info.email, info.address]
    connection.query(query, values, function(error, results, fields) {
        if (error) throw error
            // 成功后的操作，例如输出结果
        console.log('Row inserted: ', results)
        res.send('ok')
    })
})

app.get('/contacts', (req, res) => {
    connection.query('SELECT * FROM contacts', (error, results, fields) => {
        if (error) throw error
            // 发送数据给前端
        res.json(results)
    })
})

app.post('/del', (req, res) => {
    let id = req.body.id
    const deleteSql = `DELETE FROM contacts WHERE id = ?`

    // 执行删除操作
    connection.query(deleteSql, [id], (error, results, fields) => {
        if (error) {
            // 处理错误
            console.error('删除数据时出错:', error.message)
            return
        }

        // 删除成功
        console.log('数据删除成功，影响行数:', results.affectedRows)
        res.send('ok')
    })
})

app.post('/edit', (req, res) => {
    let info = req.body
    let newp
    console.log(info.newphone)
    connection.query(
        'SELECT * FROM contacts where id = ' + info.id,
        (error, results, fields) => {
            if (error) throw error
                // 发送数据给前端
            console.log(results[0].phone)
            if (results[0].phone == null) {
                results[0].phone = 'none'
            }
            newp = results[0].phone
            console.log(results[0].phone.split(';'))
            if (results[0].phone.split(';')[0] != info.phone) {
                newp = info.phone
                for (let i = 1; i < results[0].phone.split(';'); i++) {
                    newp += ';' + results[0].phone.split(';')[i]
                }
            }
            if (info.newphone && results[0].phone.split(';')[0] == info.phone) {
                newp = results[0].phone + ';' + info.newphone
            } else if (info.newphone && results[0].phone.split(';')[0] != info.phone) {
                newp += ';' + info.newphone
            }
            const updateQuery =
                'UPDATE contacts SET name = ?, phone = ?, email = ?, address = ? WHERE id = ?'
            newData = [info.name, newp, info.email, info.address, info.id]
            connection.query(updateQuery, newData, function(error, results, fields) {
                if (error) throw error
                    // 更新成功的处理逻辑
                console.log('Row updated successfully')
                res.send('ok')
            })
        }
    )
})

app.post('/output', (req, res) => {
    let data = req.body
    console.log(data)
        // 创建工作簿
    const workbook = XLSX.utils.book_new();
    let xlsdata = [
        ['id', '姓名', '电话', '邮箱', '地址']
    ]
    data.forEach(item => {
            xlsdata.push([item.id, item.name, item.phone, item.email, item.address])
        })
        // 将数据转换为工作表
    const worksheet = XLSX.utils.aoa_to_sheet(xlsdata);
    // 将工作表添加到工作簿
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    // 写入文件
    XLSX.writeFile(workbook, "./pages/output.xlsx");
    res.send('ok')
})

app.post('/input', (req, res) => {
    let data = req.body
    console.log(data)
    data.forEach(item => {
        const query =
            'INSERT INTO contacts (name, phone, email, address) VALUES (?, ?, ?, ?)'
        let values = [item['姓名'], item['电话'], item['邮箱'], item['地址']]
        connection.query(query, values, function(error, results, fields) {
            if (error) throw error
                // 成功后的操作，例如输出结果
            console.log('Row inserted: ', results)
        })
    })
    res.send('ok')
})

app.listen(5000, () => {
    console.log('服务在http://127.0.0.1:5000/ 启动')
})