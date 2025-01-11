import fs from 'node:fs';
import path from 'node:path';
import url from 'node:url';
// 获取当前脚本路径
const __filename = url.fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// 输入 CSV 文件路径
const csvFilePath = path.resolve(__dirname, 'init.csv');
// 输出 TXT 文件路径
const txtFilePath = path.resolve(__dirname, 'extraction_ip.txt');

// 提取列
const ipcom = 'ip'
const portcom = 'port'
async function extractIpAndPort() {
  try {
    // 读取 CSV 文件内容
    const data = await fs.promises.readFile(csvFilePath, 'utf8');
    // 按行分割 CSV 内容
    const lines = data.split('\n').filter(line => line.trim()); // 去掉空行
    if (lines.length < 2) {
      throw new Error('CSV 文件内容不足或格式不正确');
    }

    // 获取表头
    const headers = lines[0].split(',');
    const ipIndex = headers.indexOf(ipcom);
    const portIndex = headers.indexOf(portcom);

    if (ipIndex === -1 || portIndex === -1) {
      throw new Error(`CSV 文件缺少 ${ipcom} 或 ${portcom} 列`);
    }

    // 提取 IP 和端口
    const result = Array.from(new Set(
      lines.slice(1) // 去掉表头
        .map(line => line.split(',')) // 按逗号分割每一行
        .filter(fields => fields.length > Math.max(ipIndex, portIndex)) // 确保有足够的列
        .map(fields => {
          const ip = fields[ipIndex];
          const port = fields[portIndex];
          return `${ip} ${port}`;
        })
    )).join('\n'); // 合并成多行字符串
    // 写入到 TXT 文件
    await fs.promises.writeFile(txtFilePath, result, 'utf8');
    console.log(`已成功提取到 ${txtFilePath}`);
  } catch (error) {
    console.error('处理文件时发生错误:', error.message);
  }
}

await extractIpAndPort();
