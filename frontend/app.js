const API = 'http://localhost:8000';
let selectedSrcFiles = [];
let selectedTplFile = null;

// 拖拽支持
function setupDrop(zoneId, onFiles) {
  const zone = document.getElementById(zoneId);
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.style.borderColor = '#1677ff'; });
  zone.addEventListener('dragleave', () => { zone.style.borderColor = '#d9d9d9'; });
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.style.borderColor = '#d9d9d9';
    onFiles(e.dataTransfer.files);
  });
}

setupDrop('srcZone', files => {
  document.getElementById('srcInput').files = files;
  onSrcSelected(files);
});
setupDrop('tplZone', files => {
  document.getElementById('tplInput').files = files;
  onTplSelected(files);
});

document.getElementById('srcInput').addEventListener('change', e => onSrcSelected(e.target.files));
document.getElementById('tplInput').addEventListener('change', e => onTplSelected(e.target.files));

function onSrcSelected(files) {
  selectedSrcFiles = Array.from(files);
  const list = document.getElementById('srcList');
  list.innerHTML = selectedSrcFiles.map(f =>
    `<div class="file-item"><span class="file-name">📄 ${f.name}</span><span class="file-size">${(f.size/1024).toFixed(1)} KB</span><span class="tag tag-wait">待上传</span></div>`
  ).join('');
  document.getElementById('uploadBtn').disabled = selectedSrcFiles.length === 0;
  document.getElementById('srcMsg').innerHTML = '';
}

function onTplSelected(files) {
  if (!files.length) return;
  selectedTplFile = files[0];
  document.getElementById('tplInfo').innerHTML =
    `<div class="file-item"><span class="file-name">📋 ${selectedTplFile.name}</span><span class="file-size">${(selectedTplFile.size/1024).toFixed(1)} KB</span><span class="tag tag-wait">已选择</span></div>`;
  document.getElementById('fillBtn').disabled = false;
}

async function uploadSources() {
  if (!selectedSrcFiles.length) return;
  const btn = document.getElementById('uploadBtn');
  btn.disabled = true;
  btn.textContent = '上传中...';
  document.getElementById('srcMsg').innerHTML = '<div class="msg msg-info">正在提取文档内容...</div>';

  const fd = new FormData();
  selectedSrcFiles.forEach(f => fd.append('files', f));

  try {
    const res = await fetch(`${API}/api/upload-sources`, { method: 'POST', body: fd });
    const data = await res.json();

    if (!res.ok || !data.details) {
      throw new Error(data.detail || data.message || `服务器错误 ${res.status}`);
    }

    const list = document.getElementById('srcList');
    list.innerHTML = data.details.map(d =>
      `<div class="file-item">
        <span class="file-name">📄 ${d.filename}</span>
        <span class="file-size">${d.chars ? d.chars.toLocaleString() + ' 字符' : ''}</span>
        <span class="tag ${d.status === '成功' ? 'tag-ok' : 'tag-err'}">${d.status}</span>
      </div>`
    ).join('');

    document.getElementById('srcCnt').textContent = `${data.total_sources} 个`;
    document.getElementById('srcMsg').innerHTML =
      `<div class="msg msg-info">✅ 成功提取 ${data.uploaded} 个文件，共 ${data.total_sources} 个数据源已就绪</div>`;
  } catch (e) {
    document.getElementById('srcMsg').innerHTML = `<div class="msg msg-err">❌ 上传失败: ${e.message}</div>`;
  } finally {
    btn.disabled = false;
    btn.textContent = '上传并提取文本';
  }
}

async function clearSources() {
  try {
    await fetch(`${API}/api/sources`, { method: 'DELETE' });
    selectedSrcFiles = [];
    document.getElementById('srcList').innerHTML = '';
    document.getElementById('srcCnt').textContent = '0 个';
    document.getElementById('srcMsg').innerHTML = '<div class="msg msg-info">数据源已清空</div>';
    document.getElementById('uploadBtn').disabled = true;
  } catch (e) {
    alert('清空失败: ' + e.message);
  }
}

async function fillTemplate() {
  if (!selectedTplFile) return;

  const btn = document.getElementById('fillBtn');
  btn.disabled = true;
  document.getElementById('fillLoading').style.display = 'flex';
  document.getElementById('fillResult').innerHTML = '';

  const fd = new FormData();
  fd.append('template', selectedTplFile);
  const req = document.getElementById('reqInput').value.trim();
  if (req) fd.append('requirement', req);

  const start = Date.now();
  try {
    const res = await fetch(`${API}/api/fill-template`, { method: 'POST', body: fd });
    const data = await res.json();
    const elapsed = ((Date.now() - start) / 1000).toFixed(1);

    if (!res.ok) {
      document.getElementById('fillResult').innerHTML =
        `<div class="result-box"><div class="rrow"><span class="rlabel">状态</span><span class="rval err">❌ 失败</span></div><div class="rrow"><span class="rlabel">错误信息</span><span class="rval err">${data.detail}</span></div></div>`;
      return;
    }

    document.getElementById('fillResult').innerHTML = `
      <div class="result-box">
        <div class="rrow"><span class="rlabel">状态</span><span class="rval ok">✅ 填表成功</span></div>
        <div class="rrow"><span class="rlabel">响应时间</span><span class="rval">${data.elapsed_seconds} 秒</span></div>
        <div class="rrow"><span class="rlabel">输出文件</span><span class="rval">${data.output_file}</span></div>
        <a class="dl-btn" href="${API}${data.download_url}" download>⬇️ 下载填写结果</a>
      </div>`;
  } catch (e) {
    document.getElementById('fillResult').innerHTML =
      `<div class="result-box"><div class="rrow"><span class="rlabel">状态</span><span class="rval err">❌ 请求失败</span></div><div class="rrow"><span class="rlabel">错误</span><span class="rval err">${e.message}</span></div></div>`;
  } finally {
    btn.disabled = false;
    document.getElementById('fillLoading').style.display = 'none';
  }
}
