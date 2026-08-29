/* ================================================================
   美团外卖店铺菜单抓取脚本（支持虚拟滚动，逐个分类加载）
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
    var cw='<cols><col min="1" max="1" width="18"/><col min="2" max="2" width="42"/><col min="3" max="3" width="10"/><col min="4" max="4" width="12"/><col min="5" max="5" width="10"/></cols>';
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
    setTimeout(function(){ document.body.removeChild(a); URL.revokeObjectURL(a.href); }, 300);
  }

  /* ---------- 美团抓取核心逻辑 ---------- */
  var allProducts = [];
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

  // 点击分类（点击子元素）
  function clickCategory(index){
    var items = document.querySelectorAll('a.item_bpZh4h');
    if(index >= items.length) return false;
    var el = items[index];
    // 点击第一个子元素
    if(el.children.length > 0){
      el.children[0].click();
    } else {
      el.click();
    }
    return true;
  }

  // 提取当前分类的菜品（从.spu_s6NtPr元素提取）
  function extractProducts(category){
    var products = [];
    var seen = {};

    // 从.spu_s6NtPr元素提取（美团菜品卡片）
    var spuItems = document.querySelectorAll('.spu_s6NtPr');
    spuItems.forEach(function(el){
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

    // 补充：从.spuListWrapper_L2kZtu提取（如果.spu_s6NtPr没有）
    if(products.length === 0){
      var wrapperItems = document.querySelectorAll('.spuListWrapper_L2kZtu');
      wrapperItems.forEach(function(el){
        var text = (el.innerText||'').trim();
        if(!text) return;
        var lines = text.split('\n').filter(function(t){ return t.trim(); });
        if(lines.length < 3) return;

        // 跳过前两行（分类名）
        var name = esc(lines[2] || lines[1]);
        if(!name || name.length < 2) return;

        var price = 0;
        var sales = '';
        var img = '';
        var imgEl = el.querySelector('img');
        if(imgEl) img = imgEl.src || '';

        lines.forEach(function(line){
          var m = line.match(/[¥￥]([0-9.]+)/);
          if(m && !price) price = parseFloat(m[1]);
          if(/月售|月销|已售|销量/.test(line)) sales = esc(line);
        });

        if(name && price > 0){
          var key = category + '|' + name;
          if(!seen[key]){
            seen[key] = 1;
            products.push({
              group: category,
              name: name,
              price: price,
              sales: sales,
              spec: '',
              img: img
            });
          }
        }
      });
    }

    return products;
  }

  // 显示进度
  function showProgress(current, total, category, productCount){
    var existing = document.getElementById('mt-scraper-progress');
    if(!existing){
      var div = document.createElement('div');
      div.id = 'mt-scraper-progress';
      div.style.cssText = 'position:fixed;top:20px;left:50%;transform:translateX(-50%);background:#165DFF;color:#fff;padding:12px 24px;border-radius:8px;z-index:999999;font-size:14px;box-shadow:0 4px 12px rgba(0,0,0,0.3);';
      document.body.appendChild(div);
      existing = div;
    }
    existing.textContent = '美团菜单抓取中 (' + current + '/' + total + ') - ' + category + ' - 已收集 ' + productCount + ' 个菜品';
  }

  // 主抓取流程
  async function scrape(){
    // 检查是否是美团店铺页面
    if(!/waimai\.meituan\.com/.test(window.location.href)){
      alert('请在美团外卖店铺页面执行此脚本');
      return;
    }

    var categories = getCategories();
    if(categories.length === 0){
      alert('未找到分类，请确认已进入店铺菜单页面');
      return;
    }

    console.log('找到 ' + categories.length + ' 个分类:', categories);

    // 逐个分类抓取
    for(var i=0;i<categories.length;i++){
      var category = categories[i];
      currentCategory = category;

      // 点击分类
      clickCategory(i);
      await sleep(2000); // 等待菜品加载

      // 滚动菜品区域加载更多
      var menuContainer = document.getElementById('spu-list-dhxu28d');
      if(menuContainer){
        menuContainer.scrollTop = 0;
        await sleep(300);
        for(var s=0;s<5;s++){
          menuContainer.scrollTop = menuContainer.scrollHeight;
          await sleep(400);
        }
        menuContainer.scrollTop = 0;
        await sleep(300);
      }

      // 提取菜品
      var products = extractProducts(category);
      console.log('分类 [' + category + '] 提取到 ' + products.length + ' 个菜品');

      // 合并到总列表
      products.forEach(function(p){
        var key = p.group + '|' + p.name;
        var exists = allProducts.find(function(x){ return (x.group+'|'+x.name) === key; });
        if(!exists){
          allProducts.push(p);
        }
      });

      // 显示进度
      showProgress(i+1, categories.length, category, allProducts.length);
    }

    // 完成
    var progress = document.getElementById('mt-scraper-progress');
    if(progress) progress.remove();

    if(allProducts.length === 0){
      alert('未抓取到菜品，请确认已进入店铺菜单页面');
      return;
    }

    console.log('抓取完成，共 ' + allProducts.length + ' 个菜品');

    // 生成Excel
    var rows = [['分组','菜品名称','价格(元)','月售','规格','有无图片']];
    allProducts.forEach(function(p){
      rows.push([p.group, p.name, p.price, p.sales || '', p.spec || '', p.img ? '有' : '无']);
    });

    var blob = new Blob([xlsx(rows)], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var shopName = document.title || '美团店铺';
    downloadBlob(blob, '美团菜单_' + shopName + '_' + Date.now() + '.xlsx');

    alert('抓取完成：' + allProducts.length + ' 个菜品，已下载 Excel。');

    // 返回数据供外部使用
    return allProducts;
  }

  // 启动
  scrape();

})();
