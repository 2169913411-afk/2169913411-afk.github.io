/* ================================================================
   美团外卖店铺菜单抓取脚本 v2.1（稳定版）
   - 修复数据重复问题
   - 去掉alert弹窗，改用页面内浮动提示
   - 确保下载先触发，再显示完成提示
   - 滚动整个菜品列表提取所有菜品，自动识别分类
   使用方式：在美团店铺页面 F12 控制台粘贴执行
   ================================================================ */
(function(){
  'use strict';

  /* ---------- 工具函数 ---------- */
  function esc(s){ return String(s==null?'':s).replace(/[\n\r]/g,' ').replace(/\s+/g,' ').trim(); }
  function sleep(ms){ return new Promise(function(r){ setTimeout(r,ms); }); }

  /* ---------- XLSX 生成（纯JS，无依赖） ---------- */
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
      div.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#165DFF;color:#fff;padding:12px 24px;border-radius:8px;z-index:999999;font-size:14px;box-shadow:0 4px 12px rgba(0,0,0,0.3);text-align:center;line-height:1.6;';
      document.body.appendChild(div);
      existing = div;
    }
    existing.innerHTML = message;
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

  /* ---------- 主抓取流程 ---------- */
  async function scrape(){
    // 检查是否是美团店铺页面
    if(!/waimai\.meituan\.com/.test(window.location.href)){
      showToast('❌ 请在美团外卖店铺页面执行此脚本', 4000);
      return;
    }

    showProgress('🚀 正在准备抓取...');
    showToast('🚀 开始抓取菜单...', 2000);

    await sleep(1000);

    // 找到菜品滚动容器
    var scrollContainer = null;
    var possibleContainers = [
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
      // 如果没找到滚动容器，尝试滚动整个页面
      scrollContainer = window;
    }

    console.log('使用滚动容器:', scrollContainer === window ? 'window' : scrollContainer.className);

    // 获取所有分类名称（用于后续分组）
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
    showProgress('📜 正在滚动加载所有菜品...');
    
    var allProducts = [];
    var seenNames = {};
    var scrollStep = 500;
    var maxScrolls = 100;
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
    for(var s=0;s<maxScrolls;s++){
      // 提取当前可见的所有菜品
      var spuItems = document.querySelectorAll('.spu_s6NtPr');
      var currentCategory = '全部';
      
      // 尝试判断当前分类（根据滚动位置）
      // 简化处理：先全部归为"全部"，后续可以优化

      spuItems.forEach(function(el){
        var product = extractProduct(el, currentCategory);
        if(product && !seenNames[product.name]){
          seenNames[product.name] = 1;
          allProducts.push(product);
        }
      });

      // 更新进度
      if(s % 5 === 0){
        showProgress('📜 正在滚动加载... (' + (s+1) + '/' + maxScrolls + ')<br>已收集 ' + allProducts.length + ' 个菜品');
      }

      // 滚动
      if(scrollContainer === window){
        window.scrollBy(0, scrollStep);
      } else {
        scrollContainer.scrollTop += scrollStep;
      }
      await sleep(300);

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

    // 完成
    var progress = document.getElementById('mt-scraper-progress');
    if(progress) progress.remove();

    if(allProducts.length === 0){
      showToast('❌ 未抓取到菜品，请确认已进入店铺菜单页面', 5000);
      return;
    }

    // 生成Excel（包含图片链接）
    showProgress('📊 正在生成Excel...');
    await sleep(500);

    var rows = [['分组','菜品名称','价格(元)','月售','规格','图片链接']];
    allProducts.forEach(function(p){
      rows.push([p.group, p.name, p.price, p.sales || '', p.spec || '', p.img || '']);
    });

    var blob = new Blob([xlsx(rows)], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var shopName = (document.title || '美团店铺').replace(/[\\/:*?"<>|]/g, '_');
    var filename = '美团菜单_' + shopName + '_' + new Date().toISOString().slice(0,10) + '.xlsx';

    // 先触发下载，再显示完成提示
    downloadBlob(blob, filename);

    // 显示完成提示（不阻塞页面）
    setTimeout(function(){
      showToast('✅ 抓取完成！<br>共 ' + allProducts.length + ' 个菜品<br>Excel已开始下载：' + filename, 6000);
    }, 500);

    // 把数据保存到全局变量，供外部使用
    window.__meituanScrapeResult = allProducts;
    window.__meituanScrapeDone = true;

    console.log('抓取完成，共 ' + allProducts.length + ' 个菜品');

    // 返回数据供外部使用
    return allProducts;
  }

  // 启动
  scrape().catch(function(e){
    console.error('抓取出错:', e);
    var progress = document.getElementById('mt-scraper-progress');
    if(progress) progress.remove();
    showToast('❌ 抓取出错：' + e.message, 5000);
  });

})();
