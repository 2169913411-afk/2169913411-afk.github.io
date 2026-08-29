/* ================================================================
   美团外卖店铺菜单+图片抓取脚本 v3.0
   - 支持菜单抓取（生成Excel）
   - 支持图片抓取（打包成ZIP）
   - 支持两者同时抓取
   - 修复数据重复问题
   - 去掉alert弹窗，改用页面内浮动提示
   - 确保下载先触发，再显示完成提示
   使用方式：在美团店铺页面 F12 控制台粘贴执行
   ================================================================ */
(function(){
  'use strict';

  /* ---------- 配置 ---------- */
  var CONFIG = {
    maxImages: 200,        // 最大下载图片数量
    scrollStep: 150,       // 滚动步长（减小步长，更细致地滚动）
    maxScrolls: 120,       // 最大滚动次数
    scrollWait: 800,       // 滚动后等待时间（增加等待时间，让页面渲染内容）
    imageTimeout: 10000    // 单张图片下载超时(ms)
  };

  /* ---------- 工具函数 ---------- */
  function esc(s){ return String(s==null?'':s).replace(/[\n\r]/g,' ').replace(/\s+/g,' ').trim(); }
  function escFileName(s){ return String(s==null?'':s).replace(/[\\/:*?"<>|]/g,'_').replace(/\s+/g,' ').trim(); }
  function sleep(ms){ return new Promise(function(r){ setTimeout(r,ms); }); }

  /* ---------- ZIP/XLSX 生成（纯JS，无依赖） ---------- */
  function u8(s){ return new TextEncoder().encode(s); }
  function crc(b){
    var c,t=[],k,n;
    for(n=0;n<256;n++){ c=n; for(k=0;k<8;k++) c=(c&1)?0xEDB88320^(c>>>1):(c>>>1); t[n]=c>>>0; }
    var r=0xFFFFFFFF;
    for(var i=0;i<b.length;i++) r=t[(r^b[i])&255]^(r>>>8);
    return (r^0xFFFFFFFF)>>>0;
  }
  function zip(fs){
    var ch=[],ce=[],off=0,i;
    for(i=0;i<fs.length;i++){
      var nm=u8(fs[i].name),dt=fs[i].data,cr=crc(dt);
      var lh=new DataView(new ArrayBuffer(30));
      lh.setUint32(0,0x04034b50,true); lh.setUint16(4,20,true); lh.setUint16(6,0x0800,true);
      lh.setUint32(14,cr,true); lh.setUint32(18,dt.length,true); lh.setUint32(22,dt.length,true);
      lh.setUint16(26,nm.length,true);
      ch.push(new Uint8Array(lh.buffer),nm,dt);
      var cd=new DataView(new ArrayBuffer(46));
      cd.setUint32(0,0x02014b50,true); cd.setUint16(4,20,true); cd.setUint16(6,20,true);
      cd.setUint16(8,0x0800,true); cd.setUint32(16,cr,true); cd.setUint32(20,dt.length,true);
      cd.setUint32(24,dt.length,true); cd.setUint16(28,nm.length,true); cd.setUint32(42,off,true);
      ce.push(new Uint8Array(cd.buffer),nm); off+=30+nm.length+dt.length;
    }
    var co=off,cs=0;
    for(i=0;i<ce.length;i++) cs+=ce[i].length;
    var eo=new DataView(new ArrayBuffer(22));
    eo.setUint32(0,0x06054b50,true); eo.setUint16(8,fs.length,true); eo.setUint16(10,fs.length,true);
    eo.setUint32(12,cs,true); eo.setUint32(16,co,true);
    var all=ch.concat(ce,[new Uint8Array(eo.buffer)]),tot=0,p=0;
    for(i=0;i<all.length;i++) tot+=all[i].length;
    var out=new Uint8Array(tot);
    for(i=0;i<all.length;i++){ out.set(all[i],p); p+=all[i].length; }
    return out;
  }
  function xmlesc(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function xmlRows(rows){
    var sd='<sheetData>';
    for(var i=0;i<rows.length;i++){
      sd+='<row r="'+(i+1)+'">';
      for(var j=0;j<rows[i].length;j++){
        var ref=String.fromCharCode(65+j)+(i+1);
        var v=rows[i][j];
        if(typeof v==='number') sd+='<c r="'+ref+'" t="n"><v>'+v+'</v></c>';
        else sd+='<c r="'+ref+'" t="inlineStr"><is><t>'+xmlesc(v)+'</t></is></c>';
      }
      sd+='</row>';
    }
    return sd+'</sheetData>';
  }
  function xlsx(rows){
    var cw='<cols><col min="1" max="1" width="18"/><col min="2" max="2" width="42"/><col min="3" max="3" width="10"/><col min="4" max="4" width="12"/><col min="5" max="5" width="10"/><col min="6" max="6" width="60"/></cols>';
    var sd=xmlRows(rows);
    var ct='<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';
    var rl='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var wb='<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="菜单" sheetId="1" r:id="rId1"/></sheets></workbook>';
    var wbr='<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';
    return zip([
      {name:'[Content_Types].xml',data:u8(ct)},
      {name:'_rels/.rels',data:u8(rl)},
      {name:'xl/workbook.xml',data:u8(wb)},
      {name:'xl/_rels/workbook.xml.rels',data:u8(wbr)},
      {name:'xl/worksheets/sheet1.xml',data:u8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+cw+sd+'</worksheet>')}
    ]);
  }
  function downloadBlob(blob, filename){
    var a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download=filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function(){ document.body.removeChild(a); URL.revokeObjectURL(a.href); }, 500);
  }

  /* ---------- 页面内浮动提示（替代alert） ---------- */
  function showToast(message, duration){
    duration = duration || 3000;
    var existing = document.getElementById('mt-scraper-toast');
    if(existing) existing.remove();

    var div = document.createElement('div');
    div.id = 'mt-scraper-toast';
    div.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);background:rgba(0,0,0,0.85);color:#fff;padding:16px 32px;border-radius:8px;z-index:9999999;font-size:15px;box-shadow:0 4px 20px rgba(0,0,0,0.3);text-align:center;max-width:80%;line-height:1.6;';
    div.innerHTML = message;
    document.body.appendChild(div);

    setTimeout(function(){
      if(div.parentNode) div.parentNode.removeChild(div);
    }, duration);
  }

  /* ---------- 显示进度 ---------- */
  function showProgress(message){
    var existing = document.getElementById('mt-scraper-progress');
    if(!existing){
      var div = document.createElement('div');
      div.id = 'mt-scraper-progress';
      div.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#165DFF;color:#fff;padding:12px 24px;border-radius:8px;z-index:999999;font-size:14px;box-shadow:0 4px 12px rgba(0,0,0,0.3);text-align:center;line-height:1.6;max-width:80%;';
      document.body.appendChild(div);
      existing = div;
    }
    existing.innerHTML = message;
  }

  /* ---------- 显示选择对话框 ---------- */
  function showScrapeTypeDialog(callback){
    var existing = document.getElementById('mt-scraper-dialog');
    if(existing) existing.remove();

    var overlay = document.createElement('div');
    overlay.id = 'mt-scraper-dialog';
    overlay.style.cssText = 'position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.5);z-index:9999998;display:flex;align-items:center;justify-content:center;';

    var dialog = document.createElement('div');
    dialog.style.cssText = 'background:#fff;border-radius:12px;padding:24px;max-width:400px;width:90%;box-shadow:0 8px 32px rgba(0,0,0,0.2);';
    dialog.innerHTML = `
      <h3 style="margin:0 0 16px 0;font-size:18px;color:#1D2129;">选择抓取方式</h3>
      <div style="display:flex;flex-direction:column;gap:12px;">
        <label style="display:flex;align-items:center;gap:10px;cursor:pointer;padding:12px;border:2px solid #165DFF;background:#F0F5FF;border-radius:8px;transition:all 0.2s;" onmouseover="this.style.borderColor='#0E42D2'" onmouseout="this.style.borderColor='#165DFF'">
          <input type="radio" name="scrapeType" value="monitor" checked style="width:18px;height:18px;accent-color:#165DFF;">
          <div>
            <div style="font-weight:600;color:#1D2129;">后台监听模式（推荐）⭐</div>
            <div style="font-size:12px;color:#86909C;">用户正常浏览切换分类，脚本后台自动收集所有菜品，避免403错误</div>
          </div>
        </label>
        <label style="display:flex;align-items:center;gap:10px;cursor:pointer;padding:12px;border:1px solid #E5E6EB;border-radius:8px;transition:all 0.2s;" onmouseover="this.style.borderColor='#165DFF'" onmouseout="this.style.borderColor='#E5E6EB'">
          <input type="radio" name="scrapeType" value="menu" style="width:18px;height:18px;accent-color:#165DFF;">
          <div>
            <div style="font-weight:600;color:#1D2129;">仅抓取菜单（当前分类）</div>
            <div style="font-size:12px;color:#86909C;">立即抓取当前可见分类的菜品，生成Excel表格</div>
          </div>
        </label>
        <label style="display:flex;align-items:center;gap:10px;cursor:pointer;padding:12px;border:1px solid #E5E6EB;border-radius:8px;transition:all 0.2s;" onmouseover="this.style.borderColor='#165DFF'" onmouseout="this.style.borderColor='#E5E6EB'">
          <input type="radio" name="scrapeType" value="images" style="width:18px;height:18px;accent-color:#165DFF;">
          <div>
            <div style="font-weight:600;color:#1D2129;">仅抓取图片（当前分类）</div>
            <div style="font-size:12px;color:#86909C;">立即抓取当前可见分类的菜品图片，打包成ZIP文件</div>
          </div>
        </label>
        <label style="display:flex;align-items:center;gap:10px;cursor:pointer;padding:12px;border:1px solid #E5E6EB;border-radius:8px;transition:all 0.2s;" onmouseover="this.style.borderColor='#165DFF'" onmouseout="this.style.borderColor='#E5E6EB'">
          <input type="radio" name="scrapeType" value="both" style="width:18px;height:18px;accent-color:#165DFF;">
          <div>
            <div style="font-weight:600;color:#1D2129;">菜单+图片（当前分类）</div>
            <div style="font-size:12px;color:#86909C;">立即抓取当前可见分类的菜单和图片</div>
          </div>
        </label>
      </div>
      <div style="margin-top:16px;padding:12px;background:#FFF7E6;border-radius:8px;font-size:12px;color:#D46B08;line-height:1.6;">
        💡 <strong>推荐使用后台监听模式</strong>：由于美团反爬机制严格，自动切换分类会触发403错误。使用监听模式，您只需正常点击浏览各个分类，脚本会在后台自动收集所有菜品数据。
      </div>
      <div style="display:flex;gap:12px;margin-top:20px;">
        <button id="mt-scraper-cancel" style="flex:1;padding:10px;border:1px solid #E5E6EB;border-radius:8px;background:#fff;color:#4E5969;cursor:pointer;font-size:14px;transition:all 0.2s;" onmouseover="this.style.background='#F7F8FA'" onmouseout="this.style.background='#fff'">取消</button>
        <button id="mt-scraper-confirm" style="flex:1;padding:10px;border:none;border-radius:8px;background:#165DFF;color:#fff;cursor:pointer;font-size:14px;font-weight:600;transition:all 0.2s;" onmouseover="this.style.background='#0E42D2'" onmouseout="this.style.background='#165DFF'">开始</button>
      </div>
    `;

    overlay.appendChild(dialog);
    document.body.appendChild(overlay);

    document.getElementById('mt-scraper-cancel').onclick = function(){
      overlay.remove();
      callback(null);
    };

    document.getElementById('mt-scraper-confirm').onclick = function(){
      var selected = document.querySelector('input[name="scrapeType"]:checked');
      overlay.remove();
      callback(selected ? selected.value : 'menu');
    };
  }

  /* ---------- 提取单个菜品卡片信息 ---------- */
  function extractProduct(el, category){
    var text = (el.innerText||'').trim();
    if(!text) return null;
    var lines = text.split('\n').filter(function(t){ return t.trim(); });
    if(lines.length < 2) return null;

    var name = esc(lines[0]);
    if(!name || name.length < 2) return null;

    var price = 0;
    var sales = '';
    var spec = '';
    var img = '';
    var imgEl = el.querySelector('img');
    if(imgEl) img = imgEl.src || '';

    for(var j=1;j<lines.length;j++){
      var line = lines[j];
      var m = line.match(/[¥￥]([0-9.]+)/);
      if(m && !price) price = parseFloat(m[1]);
      if(/月售|月销|已售|销量/.test(line)) sales = esc(line);
      if(!spec && (line.indexOf('人份')>-1 || line.indexOf('份')>-1) && line.length < 20) spec = esc(line);
    }

    if(name && price > 0){
      return {
        group: category || '全部',
        name: name,
        price: price,
        sales: sales,
        spec: spec,
        img: img
      };
    }
    return null;
  }

  /* ---------- 滚动收集所有菜品 ---------- */
  async function collectProducts(){
    // 找到菜品滚动容器（优先使用#scrollArea，这是美团H5的真正滚动容器）
    var scrollContainer = null;
    var possibleContainers = [
      document.getElementById('scrollArea'),
      document.querySelector('.scrollArea_gqUOEn'),
      document.getElementById('spu-list-dhxu28d'),
      document.querySelector('.spuListWrapper_L2kZtu'),
      document.querySelector('[class*=spu-list]'),
      document.querySelector('[class*=spuList]')
    ];

    for(var i=0;i<possibleContainers.length;i++){
      if(possibleContainers[i] && possibleContainers[i].scrollHeight > possibleContainers[i].clientHeight){
        scrollContainer = possibleContainers[i];
        break;
      }
    }

    if(!scrollContainer){
      scrollContainer = window;
    }

    console.log('使用滚动容器:', scrollContainer === window ? 'window' : scrollContainer.className);

    // 获取所有分类名称
    var categoryItems = document.querySelectorAll('a.item_bpZh4h');
    var categoryNames = [];
    categoryItems.forEach(function(el){
      var text = esc(el.innerText);
      if(text && text.length<30 && categoryNames.indexOf(text)===-1){
        categoryNames.push(text);
      }
    });
    console.log('找到 ' + categoryNames.length + ' 个分类:', categoryNames);

    // 滚动加载所有菜品
    var allProducts = [];
    var seenNames = {};
    var lastHeight = 0;
    var sameHeightCount = 0;

    // 滚动到顶部
    if(scrollContainer === window){
      window.scrollTo(0, 0);
    } else {
      scrollContainer.scrollTop = 0;
    }
    await sleep(500);

    // 逐步滚动，收集所有菜品
    for(var s=0;s<CONFIG.maxScrolls;s++){
      // 提取当前可见的所有菜品
      var spuItems = document.querySelectorAll('.spu_s6NtPr');

      spuItems.forEach(function(el){
        var product = extractProduct(el, '全部');
        if(product && !seenNames[product.name]){
          seenNames[product.name] = 1;
          allProducts.push(product);
        }
      });

      // 更新进度
      if(s % 5 === 0){
        showProgress('📜 正在滚动加载... (' + (s+1) + '/' + CONFIG.maxScrolls + ')<br>已收集 ' + allProducts.length + ' 个菜品');
      }

      // 滚动
      if(scrollContainer === window){
        window.scrollBy(0, CONFIG.scrollStep);
      } else {
        scrollContainer.scrollTop += CONFIG.scrollStep;
      }
      await sleep(CONFIG.scrollWait);

      // 检查是否滚动到底部
      var currentHeight = scrollContainer === window ? document.body.scrollHeight : scrollContainer.scrollHeight;
      var currentScroll = scrollContainer === window ? window.scrollY : scrollContainer.scrollTop;
      var clientHeight = scrollContainer === window ? window.innerHeight : scrollContainer.clientHeight;

      if(currentScroll + clientHeight >= currentHeight - 100){
        console.log('已滚动到底部，共滚动 ' + (s+1) + ' 次');
        break;
      }

      if(currentHeight === lastHeight){
        sameHeightCount++;
        if(sameHeightCount > 10){
          console.log('内容高度不再变化，停止滚动');
          break;
        }
      } else {
        sameHeightCount = 0;
        lastHeight = currentHeight;
      }
    }

    // 滚动回顶部
    if(scrollContainer === window){
      window.scrollTo(0, 0);
    } else {
      scrollContainer.scrollTop = 0;
    }

    console.log('滚动完成，共收集 ' + allProducts.length + ' 个菜品');
    return allProducts;
  }

  /* ---------- 下载所有菜品图片 ---------- */
  async function downloadImages(products, shopName){
    var imageProducts = products.filter(function(p){ return p.img && p.img.indexOf('http') === 0; });
    var maxDownload = Math.min(imageProducts.length, CONFIG.maxImages);

    if(maxDownload === 0){
      showToast('❌ 没有找到菜品图片', 4000);
      return 0;
    }

    showProgress('🖼️ 开始下载 ' + maxDownload + ' 张菜品图片...');
    await sleep(500);

    var files = [];
    var failed = 0;

    for(var i=0;i<maxDownload;i++){
      try {
        showProgress('🖼️ 正在下载图片 (' + (i+1) + '/' + maxDownload + ')<br>' + imageProducts[i].name);

        var imgUrl = imageProducts[i].img.split('?')[0]; // 去掉参数，获取原图
        var response = await fetch(imgUrl);
        if(!response.ok) throw new Error('HTTP ' + response.status);
        var blob = await response.blob();
        var ab = await blob.arrayBuffer();

        var ext = 'jpg';
        var mm = imgUrl.match(/\.(jpe?g|png|webp|gif)$/i);
        if(mm) ext = mm[1].toLowerCase();

        files.push({
          name: String(i+1).padStart(3,'0') + '_' + escFileName(imageProducts[i].name).slice(0,40) + '.' + ext,
          data: new Uint8Array(ab)
        });
      } catch(e) {
        failed++;
        console.error('下载失败:', imageProducts[i].name, e.message);
      }
    }

    if(files.length > 0){
      showProgress('📦 正在打包 ' + files.length + ' 张图片...');
      await sleep(500);

      var zipData = zip(files);
      var blob = new Blob([zipData], {type:'application/zip'});
      var filename = '美团菜品图_' + shopName + '_' + new Date().toISOString().slice(0,10) + '.zip';
      downloadBlob(blob, filename);

      showToast('✅ 图片下载完成！<br>共 ' + imageProducts.length + ' 个菜品，成功 ' + files.length + ' 张，失败 ' + failed + ' 张<br>ZIP已开始下载：' + filename, 6000);
      return files.length;
    } else {
      showToast('❌ 没有成功下载任何图片', 4000);
      return 0;
    }
  }

  /* ---------- 生成Excel菜单 ---------- */
  function generateExcel(products, shopName){
    if(products.length === 0){
      showToast('❌ 未抓取到菜品', 4000);
      return false;
    }

    showProgress('📊 正在生成Excel...');

    var rows = [['分组','菜品名称','价格(元)','月售','规格','图片链接']];
    products.forEach(function(p){
      rows.push([p.group, p.name, p.price, p.sales || '', p.spec || '', p.img || '']);
    });

    var blob = new Blob([xlsx(rows)], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var filename = '美团菜单_' + shopName + '_' + new Date().toISOString().slice(0,10) + '.xlsx';
    downloadBlob(blob, filename);

    showToast('✅ 菜单抓取完成！<br>共 ' + products.length + ' 个菜品<br>Excel已开始下载：' + filename, 6000);
    return true;
  }

  /* ---------- 后台监听模式（推荐） ---------- */
  function startBackgroundMonitor(){
    // 检查是否是美团店铺页面
    if(!/waimai\.meituan\.com/.test(window.location.href)){
      showToast('❌ 请在美团外卖店铺页面执行此脚本', 4000);
      return;
    }

    // 初始化数据存储
    if(!window.__monitorData){
      window.__monitorData = {
        products: [],
        seenNames: {},
        categories: {},
        startTime: new Date().toISOString(),
        observer: null
      };
    }

    // 如果已经在监听，提示用户
    if(window.__monitorData.observer){
      showToast('📡 后台监听已在运行中！<br>当前已收集 ' + window.__monitorData.products.length + ' 个菜品', 3000);
      updateMonitorPanel();
      return;
    }

    showToast('📡 后台监听模式已启动！<br><br>请正常点击浏览各个分类，脚本会自动收集菜品数据。<br>收集完成后点击浮动面板的"导出"按钮。', 5000);

    // 创建浮动控制面板
    createMonitorPanel();

    // 立即收集当前可见的菜品
    collectCurrentDishes();

    // 启动MutationObserver监听DOM变化
    var observer = new MutationObserver(function(mutations){
      // 延迟处理，避免频繁触发
      clearTimeout(window.__monitorCollectTimer);
      window.__monitorCollectTimer = setTimeout(function(){
        collectCurrentDishes();
      }, 500);
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
      characterData: false
    });

    window.__monitorData.observer = observer;
    console.log('后台监听模式已启动，MutationObserver已激活');
  }

  /* ---------- 创建浮动控制面板 ---------- */
  function createMonitorPanel(){
    var existing = document.getElementById('mt-monitor-panel');
    if(existing) existing.remove();

    var panel = document.createElement('div');
    panel.id = 'mt-monitor-panel';
    panel.style.cssText = 'position:fixed;top:20px;right:20px;width:280px;background:#fff;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,0.2);z-index:9999999;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;overflow:hidden;';

    panel.innerHTML = `
      <div style="background:linear-gradient(135deg,#165DFF,#0E42D2);color:#fff;padding:14px 16px;display:flex;align-items:center;justify-content:space-between;">
        <div style="display:flex;align-items:center;gap:8px;">
          <span style="font-size:18px;">📡</span>
          <span style="font-weight:600;font-size:14px;">后台监听中</span>
        </div>
        <div id="mt-monitor-status" style="width:8px;height:8px;background:#00FF00;border-radius:50%;box-shadow:0 0 8px #00FF00;"></div>
      </div>
      <div style="padding:14px 16px;">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
          <span style="font-size:12px;color:#86909C;">已收集菜品</span>
          <span id="mt-monitor-count" style="font-size:24px;font-weight:700;color:#165DFF;">0</span>
        </div>
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">
          <span style="font-size:12px;color:#86909C;">已浏览分类</span>
          <span id="mt-monitor-categories" style="font-size:16px;font-weight:600;color:#4E5969;">0</span>
        </div>
        <div id="mt-monitor-current-cate" style="font-size:12px;color:#86909C;margin-bottom:12px;padding:8px;background:#F7F8FA;border-radius:6px;">
          当前分类：加载中...
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px;">
          <button id="mt-monitor-export-menu" style="padding:8px;border:1px solid #165DFF;border-radius:6px;background:#fff;color:#165DFF;cursor:pointer;font-size:12px;font-weight:600;transition:all 0.2s;" onmouseover="this.style.background='#F0F5FF'" onmouseout="this.style.background='#fff'">📊 导出Excel</button>
          <button id="mt-monitor-export-images" style="padding:8px;border:1px solid #FF7D00;border-radius:6px;background:#fff;color:#FF7D00;cursor:pointer;font-size:12px;font-weight:600;transition:all 0.2s;" onmouseover="this.style.background='#FFF7E6'" onmouseout="this.style.background='#fff'">🖼️ 导出图片</button>
        </div>
        <button id="mt-monitor-export-all" style="width:100%;padding:10px;border:none;border-radius:6px;background:linear-gradient(135deg,#165DFF,#0E42D2);color:#fff;cursor:pointer;font-size:13px;font-weight:600;transition:all 0.2s;margin-bottom:8px;" onmouseover="this.style.opacity='0.9'" onmouseout="this.style.opacity='1'">📦 导出全部（菜单+图片）</button>
        <button id="mt-monitor-stop" style="width:100%;padding:8px;border:1px solid #E5E6EB;border-radius:6px;background:#fff;color:#86909C;cursor:pointer;font-size:12px;transition:all 0.2s;" onmouseover="this.style.background='#F7F8FA'" onmouseout="this.style.background='#fff'">⏹️ 停止监听</button>
      </div>
      <div style="padding:10px 16px;background:#F7F8FA;border-top:1px solid #E5E6EB;font-size:11px;color:#86909C;line-height:1.5;">
        💡 请正常点击左侧分类浏览，脚本会自动收集每个分类的菜品。收集完所有分类后点击导出按钮。
      </div>
    `;

    document.body.appendChild(panel);

    // 绑定按钮事件
    document.getElementById('mt-monitor-export-menu').onclick = function(){
      exportMonitorData('menu');
    };
    document.getElementById('mt-monitor-export-images').onclick = function(){
      exportMonitorData('images');
    };
    document.getElementById('mt-monitor-export-all').onclick = function(){
      exportMonitorData('both');
    };
    document.getElementById('mt-monitor-stop').onclick = function(){
      stopMonitor();
    };
  }

  /* ---------- 更新监控面板显示 ---------- */
  function updateMonitorPanel(){
    if(!window.__monitorData) return;

    var countEl = document.getElementById('mt-monitor-count');
    var cateEl = document.getElementById('mt-monitor-categories');
    var currentCateEl = document.getElementById('mt-monitor-current-cate');

    if(countEl) countEl.textContent = window.__monitorData.products.length;
    if(cateEl) cateEl.textContent = Object.keys(window.__monitorData.categories).length;

    var activeCate = document.querySelector('.item_peLjjt.on_MX5HsQ') || document.querySelector('.category_Qve6pO .item_bpZh4h.active_vMJjpr');
    if(currentCateEl && activeCate){
      currentCateEl.innerHTML = '当前分类：<strong style="color:#1D2129;">' + activeCate.innerText.trim() + '</strong>';
    }
  }

  /* ---------- 收集当前可见的菜品 ---------- */
  function collectCurrentDishes(){
    if(!window.__monitorData) return;

    var activeCateEl = document.querySelector('.item_peLjjt.on_MX5HsQ') || document.querySelector('.category_Qve6pO .item_bpZh4h.active_vMJjpr');
    var categoryName = activeCateEl ? activeCateEl.innerText.trim() : '未知分类';

    // 记录已浏览的分类
    if(categoryName && categoryName !== '未知分类'){
      window.__monitorData.categories[categoryName] = true;
    }

    // 收集当前可见的菜品
    var spuItems = document.querySelectorAll('.spu_s6NtPr');
    var newCount = 0;

    spuItems.forEach(function(el){
      var product = extractProduct(el, categoryName);
      if(product && !window.__monitorData.seenNames[product.name]){
        window.__monitorData.seenNames[product.name] = 1;
        window.__monitorData.products.push(product);
        newCount++;
      }
    });

    if(newCount > 0){
      console.log('新增 ' + newCount + ' 个菜品，总计 ' + window.__monitorData.products.length + ' 个');
    }

    updateMonitorPanel();
  }

  /* ---------- 导出监控数据 ---------- */
  async function exportMonitorData(type){
    if(!window.__monitorData || window.__monitorData.products.length === 0){
      showToast('❌ 还没有收集到菜品数据，请先浏览各个分类', 3000);
      return;
    }

    var products = window.__monitorData.products;
    var shopName = (document.title || '美团店铺').replace(/[\\/:*?"<>|]/g, '_');

    showToast('📦 正在导出 ' + products.length + ' 个菜品...', 2000);

    if(type === 'menu' || type === 'both'){
      generateExcel(products, shopName);
      await sleep(1000);
    }

    if(type === 'images' || type === 'both'){
      await downloadImages(products, shopName);
    }

    // 保存到全局变量
    window.__meituanScrapeResult = products;
    window.__meituanScrapeDone = true;

    showToast('✅ 导出完成！共 ' + products.length + ' 个菜品', 4000);
  }

  /* ---------- 停止监听 ---------- */
  function stopMonitor(){
    if(window.__monitorData && window.__monitorData.observer){
      window.__monitorData.observer.disconnect();
      window.__monitorData.observer = null;

      var panel = document.getElementById('mt-monitor-panel');
      if(panel) panel.remove();

      showToast('⏹️ 后台监听已停止<br>共收集 ' + window.__monitorData.products.length + ' 个菜品<br>数据已保留，可再次点击书签继续监听', 5000);
    }
  }

  /* ---------- 主抓取流程 ---------- */
  async function scrape(scrapeType, appendMode){
    // 检查是否是美团店铺页面
    if(!/waimai\.meituan\.com/.test(window.location.href)){
      showToast('❌ 请在美团外卖店铺页面执行此脚本', 4000);
      return;
    }

    // 检查是否是追加模式
    var isAppend = appendMode === true || (window.__meituanScrapeResult && window.__meituanScrapeResult.length > 0 && confirm('检测到已有 ' + window.__meituanScrapeResult.length + ' 个菜品的抓取结果，是否追加到已有结果中？\\n\\n点击"确定"追加（适合手动切换分类后继续抓取）\\n点击"取消"重新开始抓取'));

    if(isAppend){
      showProgress('📝 追加模式：将新抓取的菜品添加到已有结果中...');
    } else {
      // 清空之前的抓取结果
      window.__meituanScrapeResult = null;
      window.__meituanScrapeDone = false;
      showProgress('🚀 正在准备抓取...');
    }

    showToast(isAppend ? '📝 追加模式开始抓取...' : '🚀 开始抓取...', 2000);
    await sleep(1000);

    try {
      // 收集所有菜品
      var newProducts = await collectProducts();

      // 合并已有结果和新结果（去重）
      var products = [];
      var seenNames = {};

      // 先添加已有结果
      if(isAppend && window.__meituanScrapeResult){
        window.__meituanScrapeResult.forEach(function(p){
          if(!seenNames[p.name]){
            seenNames[p.name] = 1;
            products.push(p);
          }
        });
      }

      // 再添加新结果
      newProducts.forEach(function(p){
        if(!seenNames[p.name]){
          seenNames[p.name] = 1;
          products.push(p);
        }
      });

      // 移除进度条
      var progress = document.getElementById('mt-scraper-progress');
      if(progress) progress.remove();

      if(products.length === 0){
        showToast('❌ 未抓取到菜品，请确认已进入店铺菜单页面', 5000);
        return;
      }

      var shopName = (document.title || '美团店铺').replace(/[\\/:*?"<>|]/g, '_');
      var newCount = products.length - (isAppend && window.__meituanScrapeResult ? window.__meituanScrapeResult.length : 0);

      // 根据选择执行相应操作
      if(scrapeType === 'menu' || scrapeType === 'both'){
        generateExcel(products, shopName);
        await sleep(1000);
      }

      if(scrapeType === 'images' || scrapeType === 'both'){
        await downloadImages(products, shopName);
      }

      // 把数据保存到全局变量，供外部使用
      window.__meituanScrapeResult = products;
      window.__meituanScrapeDone = true;

      console.log('抓取完成，共 ' + products.length + ' 个菜品' + (isAppend ? '（新增 ' + newCount + ' 个）' : ''));

      // 显示完成提示，告诉用户可以继续抓取其他分类
      if(products.length > 0){
        setTimeout(function(){
          showToast('✅ 抓取完成！共 ' + products.length + ' 个菜品' + (isAppend ? '（新增 ' + newCount + ' 个）' : '') + '<br><br>💡 提示：如需抓取其他分类，请手动点击左侧分类切换，然后再次点击书签继续抓取（会自动追加到已有结果中）', 8000);
        }, 1000);
      }

    } catch(e){
      console.error('抓取出错:', e);
      var progress = document.getElementById('mt-scraper-progress');
      if(progress) progress.remove();
      showToast('❌ 抓取出错：' + e.message, 5000);
    }
  }

  /* ---------- 读取URL参数，自动选择抓取类型 ---------- */
  function getUrlParam(name){
    var m=location.href.match(new RegExp('[?&#]'+name+'=([^&#]*)'));
    return m?decodeURIComponent(m[1]):'';
  }

  // 启动：优先读取URL参数，否则显示选择对话框
  var urlMode = getUrlParam('cmz_scrape');
  if(urlMode === 'monitor'){
    // 后台监听模式
    startBackgroundMonitor();
  } else if(urlMode === 'menu' || urlMode === 'img' || urlMode === 'all'){
    // 将URL参数转换为脚本使用的类型
    var scrapeType = urlMode === 'menu' ? 'menu' : (urlMode === 'img' ? 'images' : 'both');
    scrape(scrapeType);
  } else {
    // 没有URL参数，显示选择对话框
    showScrapeTypeDialog(function(scrapeType){
      if(scrapeType === 'monitor'){
        startBackgroundMonitor();
      } else if(scrapeType){
        scrape(scrapeType);
      } else {
        console.log('用户取消抓取');
      }
    });
  }

})();
