/* ================================================================
   美团外卖店铺菜单抓取脚本 v2.0（优化版）
   - 修复数据重复问题
   - 去掉alert弹窗，改用页面内浮动提示
   - 确保下载先触发，再显示完成提示
   - 增加全局去重功能
   使用方式：在美团店铺页面 F12 控制台粘贴执行
   ================================================================ */
(function(){
  'use strict';

  /* ---------- 工具函数 ---------- */
  function esc(s){ return String(s==null?'':s).replace(/[\n\r]/g,' ').replace(/\s+/g,' ').trim(); }
  function num(s){ var m=String(s||'').match(/(\d+\.?\d*)/); return m?parseFloat(m[1]):0; }
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

  /* ---------- 美团抓取核心逻辑 ---------- */
  var allProducts = [];
  var globalSeen = {}; // 全局去重
  var categoryNames = [];
  var currentCategory = '';

  // 获取所有分类
  function getCategories(){
    var items = document.querySelectorAll('a.item_bpZh4h');
    var cats = [];
    items.forEach(function(el){
      var text = esc(el.innerText);
      if(text && text.length<30 && cats.indexOf(text)===-1){
        cats.push(text);
      }
    });
    return cats;
  }

  // 点击分类（点击子元素，确保触发切换）
  function clickCategory(index){
    var items = document.querySelectorAll('a.item_bpZh4h');
    if(index >= items.length) return false;
    var el = items[index];

    // 尝试多种点击方式
    try {
      // 先尝试点击子元素
      if(el.children.length > 0){
        el.children[0].click();
      } else {
        el.click();
      }
    } catch(e) {
      // 如果click失败，尝试dispatchEvent
      try {
        var event = new MouseEvent('click', {
          'view': window,
          'bubbles': true,
          'cancelable': true
        });
        el.dispatchEvent(event);
      } catch(e2) {
        console.error('点击分类失败:', e2);
      }
    }
    return true;
  }

  // 找到菜品滚动容器
  function findScrollContainer(){
    // 尝试多种可能的容器
    var selectors = [
      '#spu-list-dhxu28d',
      '.spuListWrapper_L2kZtu',
      '[class*=spu-list]',
      '[class*=spuList]',
      '[class*=menu-list]',
      '[class*=goods-list]'
    ];

    for(var i=0;i<selectors.length;i++){
      var el = document.querySelector(selectors[i]);
      if(el && el.scrollHeight > el.clientHeight){
        return el;
      }
    }

    // 如果没找到，返回body
    return document.body;
  }

  // 滚动到分类位置（美团可能是锚点滚动，不是内容切换）
  function scrollToCategory(categoryName){
    // 尝试找到分类标题元素并滚动到视图
    var allElements = document.querySelectorAll('*');
    for(var i=0;i<allElements.length;i++){
      var el = allElements[i];
      if(el.children.length === 0 && esc(el.innerText) === categoryName){
        el.scrollIntoView({behavior: 'smooth', block: 'start'});
        return true;
      }
    }
    return false;
  }

  // 提取当前可见区域的菜品
  function extractProducts(category){
    var products = [];
    var seen = {};

    // 从.spu_s6NtPr元素提取（美团菜品卡片）
    var spuItems = document.querySelectorAll('.spu_s6NtPr');
    spuItems.forEach(function(el){
      // 只提取可见的元素
      var rect = el.getBoundingClientRect();
      if(rect.top < -100 || rect.top > window.innerHeight + 100) return;

      var text = (el.innerText||'').trim();
      if(!text) return;
      var lines = text.split('\n').filter(function(t){ return t.trim(); });
      if(lines.length < 2) return;

      // 第一行是菜品名
      var name = esc(lines[0]);
      if(!name || name.length < 2) return;

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
        var key = category + '|' + name;
        if(!seen[key]){
          seen[key] = 1;
          products.push({
            group: category,
            name: name,
            price: price,
            sales: sales,
            spec: spec,
            img: img
          });
        }
      }
    });

    // 如果可见区域提取不到，尝试提取所有（备用方案）
    if(products.length === 0){
      spuItems.forEach(function(el){
        var text = (el.innerText||'').trim();
        if(!text) return;
        var lines = text.split('\n').filter(function(t){ return t.trim(); });
        if(lines.length < 2) return;

        var name = esc(lines[0]);
        if(!name || name.length < 2) return;

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
          var key = category + '|' + name;
          if(!seen[key]){
            seen[key] = 1;
            products.push({
              group: category,
              name: name,
              price: price,
              sales: sales,
              spec: spec,
              img: img
            });
          }
        }
      });
    }

    return products;
  }

  // 显示进度
  function showProgress(current, total, category, productCount, categoryCount){
    var existing = document.getElementById('mt-scraper-progress');
    if(!existing){
      var div = document.createElement('div');
      div.id = 'mt-scraper-progress';
      div.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#165DFF;color:#fff;padding:12px 24px;border-radius:8px;z-index:999999;font-size:14px;box-shadow:0 4px 12px rgba(0,0,0,0.3);text-align:center;line-height:1.6;';
      document.body.appendChild(div);
      existing = div;
    }
    existing.innerHTML = '美团菜单抓取中 (' + current + '/' + total + ')<br>当前分类：' + category + '（' + categoryCount + '个菜品）<br>累计收集：' + productCount + ' 个菜品（已去重）';
  }

  // 主抓取流程
  async function scrape(){
    // 检查是否是美团店铺页面
    if(!/waimai\.meituan\.com/.test(window.location.href)){
      showToast('❌ 请在美团外卖店铺页面执行此脚本', 4000);
      return;
    }

    var categories = getCategories();
    if(categories.length === 0){
      showToast('❌ 未找到分类，请确认已进入店铺菜单页面', 4000);
      return;
    }

    console.log('找到 ' + categories.length + ' 个分类:', categories);
    showToast('🚀 开始抓取，共 ' + categories.length + ' 个分类', 2000);

    // 逐个分类抓取
    for(var i=0;i<categories.length;i++){
      var category = categories[i];
      currentCategory = category;

      // 点击分类
      clickCategory(i);
      await sleep(2000); // 等待菜品加载（增加等待时间）

      // 滚动到分类位置（如果是锚点滚动模式）
      scrollToCategory(category);
      await sleep(500);

      // 滚动菜品区域加载更多
      var menuContainer = findScrollContainer();
      if(menuContainer && menuContainer !== document.body){
        menuContainer.scrollTop = 0;
        await sleep(300);
        // 增加滚动次数，确保加载所有菜品
        for(var s=0;s<8;s++){
          menuContainer.scrollTop = menuContainer.scrollHeight;
          await sleep(500);
        }
        menuContainer.scrollTop = 0;
        await sleep(300);
      } else {
        // 如果没找到滚动容器，滚动整个页面
        window.scrollTo(0, 0);
        await sleep(300);
        for(var s2=0;s2<5;s2++){
          window.scrollTo(0, document.body.scrollHeight);
          await sleep(500);
        }
        window.scrollTo(0, 0);
        await sleep(300);
      }

      // 提取菜品
      var products = extractProducts(category);
      console.log('分类 [' + category + '] 提取到 ' + products.length + ' 个菜品');

      // 合并到总列表（全局去重）
      var newCount = 0;
      products.forEach(function(p){
        var key = p.name; // 按菜品名称全局去重
        if(!globalSeen[key]){
          globalSeen[key] = 1;
          allProducts.push(p);
          newCount++;
        }
      });

      console.log('分类 [' + category + '] 新增 ' + newCount + ' 个菜品，累计 ' + allProducts.length + ' 个');

      // 显示进度
      showProgress(i+1, categories.length, category, allProducts.length, products.length);
    }

    // 完成
    var progress = document.getElementById('mt-scraper-progress');
    if(progress) progress.remove();

    if(allProducts.length === 0){
      showToast('❌ 未抓取到菜品，请确认已进入店铺菜单页面', 5000);
      return;
    }

    console.log('抓取完成，共 ' + allProducts.length + ' 个菜品（已去重）');

    // 生成Excel（包含图片链接）
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
      showToast('✅ 抓取完成！<br>共 ' + allProducts.length + ' 个菜品（已去重）<br>Excel已开始下载：' + filename, 6000);
    }, 500);

    // 把数据保存到全局变量，供外部使用
    window.__meituanScrapeResult = allProducts;
    window.__meituanScrapeDone = true;

    // 返回数据供外部使用
    return allProducts;
  }

  // 启动
  scrape().catch(function(e){
    console.error('抓取出错:', e);
    showToast('❌ 抓取出错：' + e.message, 5000);
  });

})();
