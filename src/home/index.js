import React, { useState } from 'react'
import logo from '../logo.svg'
import * as XLSX from 'xlsx'
import XLSX_STYLE from 'xlsx-style'
import { Button, Input, Drawer, message } from 'antd'
import { DiffOutlined, ThunderboltOutlined } from '@ant-design/icons'
import './index.css'

const CELL_STYLE = {
  font: {
    name: '宋体',
    sz: 15,
  },
  alignment: {
    horizontal: "center", vertical: "center", wrap_text: true
  }
}

const TITLE_STYLE = {
  font: {
    name: '宋体',
    sz: 16,
    bold: true,
    color: {
      rgb: "000000"
    }
  },
  fill: {
    fgColor: {
      rgb: "9dcc70"
    }
  },
  alignment: {
    horizontal: "center", vertical: "center", wrap_text: true
  }
}

const month = new Date().getMonth() + 1

const HOME = () => {
  const [list, setList] = useState([])
  const [users, setUsers] = useState([])
  const [disable, setDisable] = useState(true)
  const [visible, setVisible] = useState(false)
  const [fileName, setFileName] = useState('')
  const [outputName, setOutputName] = useState('')

  // 处理上传表格的数据
  const handleXLSXData = event => {
    try {
      const { result } = event.target
      // 以二进制流方式读取得到整份excel表格对象
      const workbook = XLSX.read(result, { type: 'binary' })
      // 存储获取到的数据
      let sheet = workbook.Sheets['每日统计']
      const rows = new Map()
      Object.keys(sheet)
        .filter(key => Number(key.slice(1)) > 4)
        .forEach(key => {
          const rowKey = `ROW-${key.slice(1)}`
          if (!rows.has(rowKey)) {
            rows.set(rowKey, {
              [key.slice(0, 1)]: sheet[key].w
            })
          } else {
            const pre = rows.get(rowKey)
            rows.set(rowKey, Object.assign(pre, {
              [key.slice(0, 1)]: sheet[key].w
            }))
          }
        })
      const nameMap = {}
      let arr = [...rows.values()]
        .filter(item => Boolean(item.X) && Number(item.X) >= 600) // 筛选出>=10小时的天数
        .filter(item => item.J.indexOf('补卡') === -1 && item.L.indexOf('补卡') === -1) // 过滤掉含补卡的

      // 设置项目名
      const teamName = arr[0] && arr[0].B && arr[0].B.split('--') && arr[0].B.split('--')[1]
      if (teamName) {
        setOutputName(`${teamName}__${month}月餐补.xlsx`)
      }
      arr = arr
        .map(item => {
          if (!nameMap[item.A]) {
            nameMap[item.A] = item.A
            const amount = arr.filter(i => i.A === item.A).length
            return {
              姓名: item.A,
              工号: item.D,
              博彦工号: '',
              加班日期: item.G,
              加班天数: amount,
              餐补总额: amount * 13,
              TL: item.B.split('-')[0],
            }
          }
          return {
            姓名: item.A,
            工号: item.D,
            博彦工号: '',
            加班日期: item.G,
          }
        })
        .reduce((res, item, index) => {
          if (item.TL && index > 0) {
            res.push({})
            res.push({})
            res.push({})
            res.push({
              姓名: '姓名',
              工号: '工号',
              博彦工号: '博彦工号',
              加班日期: '加班日期',
              加班天数: '加班天数',
              餐补总额: '餐补总额',
              TL: 'TL',
            })
            res.push(item)
          } else {
            res.push(item)
          }
          return res
        }, [])

      // console.log(arr)
      // console.log(nameMap)
      setList(arr)
      setDisable(false)
      handleUsersInfo(nameMap)
      message.success('上传成功！')
    } catch (e) {
      setDisable(true)
      message.error('文件类型/格式不正确！请上传钉钉导出的考勤报表！')
    }
  }

  // 处理用户信息（姓名、工号）
  const handleUsersInfo = (names) => {
    let storage
    try {
      storage = JSON.parse(localStorage.getItem('users'))
    } catch (e) {}
    if (storage && storage.length) {
      setUsers(Object.keys(names).map(key => {
        const user = storage.find(s => s.name === key)
        return { name: key, jobId: user ? user.jobId : '' }
      }))
    } else {
      setUsers(Object.keys(names).map(key => {
        return { name: key, jobId: '' }
      }))
    }
  }

  const onImportExcel = e => {
    // 获取上传的文件对象
    const { files } = e.target
    if (!files[0]) return

    setFileName(files[0].name)
    // 通过FileReader对象读取文件
    const fileReader = new FileReader()
    fileReader.onload = handleXLSXData
    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0])
  }

  const onDropFile = (e) => {
    e.preventDefault()
    // 获取文件对象
    var fileList = e.dataTransfer.files
    // 检测是否是拖拽文件到页面的操作
    if (fileList.length === 0) return false
    setFileName(fileList[0].name)
    // 检测文件
    const suffix = fileList[0].name.split('.')[1]
    if (!['xlsx', 'xls'].includes(suffix)) {
      message.error('请拖入.xlsx或.xls文件！')
      setDisable(true)
      return false
    }
    const fileReader = new FileReader()
    fileReader.onload = handleXLSXData
    // 以二进制方式打开文件
    fileReader.readAsBinaryString(fileList[0])
  }

  const writeJobNumber = () => {
    setVisible(true)
  }
  const handleCancel = () => {
    setVisible(false)
  }

  // 输入博彦工号
  const onJobIdInput = (e, index) => {
    const arr = users.slice()
    arr.splice(index, 1, {
      ...arr[index],
      jobId: e.target.value
    })
    let storage
    try {
      storage = JSON.parse(localStorage.getItem('users'))
    } catch (e) {}
    // console.log("storage", storage)
    if (storage) {
      arr.forEach(item => {
        const i = storage.findIndex(s => s.name === item.name)
        if (i !== -1) {
          storage.splice(i, 1, item)
        } else {
          storage.push(item)
        }
      })
    } else {
      storage = arr
    }
    localStorage.setItem('users', JSON.stringify(storage))
    setUsers(arr)
  }

  // 导出 无样式
  const exportXLSX = () => {
    const cache = {}
    const temp = list.map(item => {
      let jobId
      if (cache[item['姓名']]) {
        jobId = cache[item['姓名']]
      } else {
        if (item['博彦工号'] === '博彦工号' || Object.keys(item).length === 0) return item
        const user = users.find(u => u.name === item['姓名'])
        jobId = (user && user.jobId) || ''
        cache[item['姓名']] = jobId
      }
      return {
        ...item,
        博彦工号: jobId,
      }
    })
    const WB = XLSX.utils.book_new();
    const WS = XLSX.utils.json_to_sheet(temp);
    WS['!cols'] = [{wpx: 100}, {wpx: 100}, {wpx: 100}, {wpx: 150}, {wpx: 100}, {wpx: 100}, {wpx: 100}];

    XLSX.utils.book_append_sheet(WB, WS, 'sheet1');
    XLSX.writeFile(WB, '团队餐补统计.xlsx');
  }

  // 导出 带样式
  // 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
  const exportXLSX2 = () => {
    const sheetName = 'sheet1'
    const cache = {}
    const temp = list.map(item => {
      let jobId
      if (cache[item['姓名']]) {
        jobId = cache[item['姓名']]
      } else {
        if (item['博彦工号'] === '博彦工号' || Object.keys(item).length === 0) return item
        const user = users.find(u => u.name === item['姓名'])
        jobId = (user && user.jobId) || ''
        cache[item['姓名']] = jobId
      }
      return {
        ...item,
        博彦工号: jobId,
      }
    })
    const workbook = {
      SheetNames: [sheetName],
      Sheets: {}
    };
    const WS = XLSX.utils.json_to_sheet(temp);
    WS['!cols'] = [{wpx: 120}, {wpx: 120}, {wpx: 120}, {wpx: 200}, {wpx: 120}, {wpx: 120}, {wpx: 120}];

    const cacheTitle = {}
    for(var key in WS) {
      if (key[0] !== '!') {
        const line = key.substr(1)
        if (WS[key].v === '姓名') {
          cacheTitle[line] = line
        }

        if (cacheTitle[line]) {
          WS[key].s = TITLE_STYLE
        } else {
          WS[key].s = CELL_STYLE
        }
      }
    }
    workbook.Sheets[sheetName] = WS

    // 生成excel的配置项
    const options = {
      bookType: 'xlsx', // 要生成的文件类型
      bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
      type: 'binary'
    };

    const wbout = XLSX_STYLE.write(workbook, options);
    const blob = new Blob([s2ab(wbout)], {type: "application/octet-stream"});
    // 字符串转ArrayBuffer
    function s2ab(s) {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i=0; i!==s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
    saveAs(blob, outputName || `xxx团队__${month}月餐补统计.xlsx`)
  }

  function saveAs (url, saveName) {
    if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // 创建blob地址
    }
    var aLink = document.createElement('a');
    aLink.href = url;
    aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
    var event;
    if (window.MouseEvent) event = new MouseEvent('click');
    else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
    }
    aLink.dispatchEvent(event);
  }


  return (
    <div className="App">
      <div className="header">
        <img src={logo} className="App-logo" alt="Design by Va" title="Design by Va" />
        <p className="title">一键生成团队餐补统计</p>
      </div>
      <div
        className="dropArea"
        onDrop={onDropFile}
        onDragOver={e => e.preventDefault()}
        onDragLeave={e => e.preventDefault()}
      >
        <DiffOutlined />
        <input className="uploadInput" type='file' accept='.xlsx, .xls' onChange={onImportExcel} />
        {fileName || '点击上传/将文件拖拽到此区域'}
      </div>
      <p>支持 .xlsx、.xls 格式的文件</p>
      <div>
        <Button type="primary" className="m-r32" danger onClick={() => window.location.reload()}>初始化</Button>
        <Button className="m-r32" ghost onClick={writeJobNumber}>填入博彦工号</Button>
        <Button type="primary" disabled={disable} onClick={exportXLSX2}><ThunderboltOutlined />导出餐补统计</Button>
      </div>
      <Drawer
        title="填入博彦工号"
        width="512"
        placement="right"
        closable={false}
        onClose={handleCancel}
        visible={visible}
      >
        {users.map((item, index) => {
          return (
            <div key={item.name} className="formItem">
              <span>{item.name}：</span>
              <Input placeholder="输入博彦工号" maxLength="10" value={item.jobId} onInput={(e) => onJobIdInput(e, index)} />
            </div>
          )
        })}
      </Drawer>
    </div>
  )
}

export default HOME
