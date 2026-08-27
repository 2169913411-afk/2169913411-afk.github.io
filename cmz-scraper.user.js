// ==UserScript==
// @name         餐谋长·商家调研抓取助手（淘宝闪购/饿了么）
// @namespace    canzhang-scraper
// @version      1.0.3
// @description  进入淘宝闪购（饿了么）店铺页后自动抓取菜单/菜品图并下载 Excel/zip；自动解析淘宝口令进店。网页点「开始抓取」后全自动，无需 F12。
// @match        https://h5.ele.me/*
// @match        https://*.ele.me/*
// @match        https://market.m.taobao.com/*
// @match        https://m.tb.cn/*
// @match        https://tb.cn/*
// @match        https://*.tb.cn/*
// @match        https://2169913411-afk.github.io/*
// @grant        GM_xmlhttpRequest
// @run-at       document-idle
// @noframes
// ==/UserScript==

(function(){
  'use strict';

  /* ---------- 工具 ---------- */
  function getParam(name, u){
    u = u || location.href;
    var m = u.match(new RegExp('[?&#]' + name + '=([^&#]*)'));
    return m ? decodeURIComponent(m[1]) : '';
  }
  function sleep(ms){ return new Promise(function(r){ setTimeout(r, ms); }); }
  function esc(s){ return String(s||'').replace(/[\\\/:*?"<>|\n\r]/g, ' ').replace(/\s+/g, ' ').trim(); }
  function numText(s){ return String(s||'').replace(/[^\d.]/g, ''); }
  function txt(item, sel){
    var el = item.querySelector(sel);
    return el ? (el.innerText || '').replace(/\s+/g, '').trim() : '';
  }
  function ts(){ var d=new Date(); function p2(x){return ('0'+x).slice(-2);} return d.getFullYear()+p2(d.getMonth()+1)+p2(d.getDate())+'_'+p2(d.getHours())+p2(d.getMinutes()); }

  /* ---------- 顶部提示横幅 ---------- */
  function banner(html, color){
    var id = 'cmz-banner';
    var b = document.getElementById(id);
    if(!b){
      b = document.createElement('div');
      b.id = id;
      b.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:2147483647;padding:10px 16px;font-size:14px;font-weight:600;font-family:-apple-system,"PingFang SC",sans-serif;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.15);';
      document.documentElement.appendChild(b);
    }
    b.style.background = color || '#EFF6FF';
    b.style.color = '#1D4ED8';
    b.style.borderBottom = '2px solid #2563EB';
    b.innerHTML = html;
  }
  function removeBanner(){
    var b = document.getElementById('cmz-banner');
    if(b) b.remove();
  }

  /* ---------- 下载 ---------- */
  function downloadBlob(blob, filename){
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(function(){ document.body.removeChild(a); URL.revokeObjectURL(url); }, 400);
  }

  /* ---------- xlsx / zip 生成（内置，无外部依赖） ---------- */
  function crc32(bytes){
    var c, table=[], k, n;
    for(n=0;n<256;n++){ c=n; for(k=0;k<8;k++) c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1); table[n]=c>>>0; }
    var crc=0xFFFFFFFF; for(var i=0;i<bytes.length;i++) crc=table[(crc^bytes[i])&0xFF]^(crc>>>8);
    return (crc^0xFFFFFFFF)>>>0;
  }
  function utf8(s){ return new TextEncoder().encode(s); }
  function escapeXml(s){ return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
  function buildZip(files){
    var chunks=[], central=[], offset=0, i;
    for(i=0;i<files.length;i++){
      var name=utf8(files[i].name), data=files[i].data;
      var crc=crc32(data), size=data.length;
      var lh=new DataView(new ArrayBuffer(30));
      lh.setUint32(0,0x04034b50,true); lh.setUint16(4,20,true); lh.setUint16(6,0x0800,true);
      lh.setUint16(8,0,true); lh.setUint16(10,0,true); lh.setUint16(12,0,true);
      lh.setUint32(14,crc,true); lh.setUint32(18,size,true); lh.setUint32(22,size,true);
      lh.setUint16(26,name.length,true); lh.setUint16(28,0,true);
      chunks.push(new Uint8Array(lh.buffer), name, data);
      var ch=new DataView(new ArrayBuffer(46));
      ch.setUint32(0,0x02014b50,true); ch.setUint16(4,20,true); ch.setUint16(6,20,true);
      ch.setUint16(8,0x0800,true); ch.setUint16(10,0,true); ch.setUint16(12,0,true); ch.setUint16(14,0,true);
      ch.setUint32(16,crc,true); ch.setUint32(20,size,true); ch.setUint32(24,size,true);
      ch.setUint16(28,name.length,true); ch.setUint16(30,0,true); ch.setUint16(32,0,true);
      ch.setUint16(34,0,true); ch.setUint16(36,0,true); ch.setUint32(38,0,true); ch.setUint32(42,offset,true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + size;
    }
    var centralOffset=offset, centralSize=0;
    for(i=0;i<central.length;i++) centralSize += central[i].length;
    var eocd=new DataView(new ArrayBuffer(22));
    eocd.setUint32(0,0x06054b50,true); eocd.setUint16(4,0,true); eocd.setUint16(6,0,true);
    eocd.setUint16(8,files.length,true); eocd.setUint16(10,files.length,true);
    eocd.setUint32(12,centralSize,true); eocd.setUint32(16,centralOffset,true); eocd.setUint16(20,0,true);
    var all=chunks.concat(central,[new Uint8Array(eocd.buffer)]);
    var total=0,p=0;
    for(i=0;i<all.length;i++) total += all[i].length;
    var out=new Uint8Array(total);
    for(i=0;i<all.length;i++){ out.set(all[i],p); p += all[i].length; }
    return out;
  }
  function buildMenuXlsx(rows){
    var width=[22,40,10,14,10,18,12];
    var colsXml='<cols>'+width.map(function(w,i){ return '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>'; }).join('')+'</cols>';
    var sheetXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'+colsXml+'<sheetData>'+
      rows.map(function(r,ri){
        var isHeader=ri===0;
        var cells=r.map(function(c,ci){
          var ref=String.fromCharCode(65+ci)+(ri+1);
          var style=isHeader?'1':'0';
          if(typeof c==='number') return '<c r="'+ref+'" s="'+style+'"><v>'+c+'</v></c>';
          return '<c r="'+ref+'" s="'+style+'" t="inlineStr"><is><t>'+escapeXml(c)+'</t></is></c>';
        }).join('');
        return '<row r="'+(ri+1)+'">'+cells+'</row>';
      }).join('')+'</sheetData></worksheet>';
    var ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
      '<Default Extension="xml" ContentType="application/xml"/>'+
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';
    var rels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var wb='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+
      '<sheets><sheet name="菜单" sheetId="1" r:id="rId1"/></sheets></workbook>';
    var wbr='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';
    return buildZip([
      {name:'[Content_Types].xml', data:utf8(ct)},
      {name:'_rels/.rels', data:utf8(rels)},
      {name:'xl/workbook.xml', data:utf8(wb)},
      {name:'xl/_rels/workbook.xml.rels', data:utf8(wbr)},
      {name:'xl/worksheets/sheet1.xml', data:utf8(sheetXml)}
    ]);
  }

  /* ---------- 跨域下载图片 ---------- */
  function gmFetch(url, type){
    return new Promise(function(res, rej){
      try{
        GM_xmlhttpRequest({
          method: 'GET', url: url, responseType: type || 'blob', timeout: 20000,
          onload: function(r){ if(r.status>=200 && r.status<300) res(r.response); else rej(new Error('http '+r.status)); },
          onerror: rej, ontimeout: rej
        });
      }catch(e){ rej(e); }
    });
  }

  /* ---------- 提取菜单 ---------- */
  function parseMenu(){
    var items=[], seen={};
    document.querySelectorAll('.menuItem').forEach(function(item){
      var cate = item.getAttribute('data-cate-name') || '';
      var title = txt(item, '.menuItem--info-title') || txt(item, '.menuItem--info-title--warp');
      var sales = txt(item, '.menuItem--info-sales');               // 月售 400+
      var price = numText(txt(item, '.menuItem--info-price'));      // 现价
      var use   = numText(txt(item, '.menuItem--info-useCouponPrice')); // 券后/预估到手
      var desc  = txt(item, '.menuItem--info-description');
      var img=''; var im = item.querySelector('.menuItem--image-img');
      if(im) img = im.getAttribute('src') || '';
      if(!title) return;
      var key = title + '|' + price;
      if(seen[key]) return; seen[key]=1;
      items.push({cate:cate, name:title, sales:sales, price:price, useCoupon:use, desc:desc, img:img});
    });
    return items;
  }

  /* ---------- 滚动加载全部菜单 ---------- */
  async function loadAll(){
    var sc = document.querySelector('.mor-comp-page-content');
    if(!sc){
      // 备选：找任意可滚动大容器
      var all=document.querySelectorAll('*');
      for(var i=0;i<all.length;i++){
        var el=all[i];
        if(el.scrollHeight>el.clientHeight+200 && el.clientHeight>100){ sc=el; break; }
      }
    }
    if(!sc) return;
    var prev=0, stable=0;
    for(var i=0;i<60 && stable<8; i++){
      sc.scrollTop = sc.scrollHeight;
      await sleep(400 + Math.floor(Math.random()*500)); // 随机间隔，模拟真人滚动，降低风控概率
      // 若被风控跳转，立即停止
      if(/punish|waimai-guide|x5secdata/i.test(location.href)) return;
      var n = document.querySelectorAll('.menuItem').length;
      if(n===prev){ stable++; } else { stable=0; prev=n; }
    }
    sc.scrollTop = 0;
  }

  /* ---------- 构造店铺 URL ---------- */
  function buildShopUrl(shopId, mode){
    var base='https://h5.ele.me/2021001185671035/pages/ele-takeout-index/ele-takeout-index';
    var p=['from=mobile.default','psc=24901','type=1','scheme_type=SHOP_SCHEME','latitude=31.222007','longitude=121.353007','shopId='+shopId,'cmz_scrape='+(mode||'menu')];
    return base+'?'+p.join('&');
  }

  /* ---------- 主流程 ---------- */
  var mode = getParam('cmz_scrape');
  var host = location.hostname;

  /* 站点自身：标记插件已安装，供网页识别 */
  if(host.indexOf('2169913411-afk.github.io') >= 0){
    try{ document.documentElement.setAttribute('data-cmz-scraper', 'installed'); }catch(e){}
    return;
  }

  /* 淘宝口令短链 m.tb.cn / tb.cn：直接抓取响应提取 shopId → 自动进入店铺并抓取（GM_xmlhttpRequest 不受 CORS 限制） */
  if(host.indexOf('tb.cn') >= 0 && host.indexOf('market.') < 0){
    banner('🔗 正在解析口令，识别店铺…');
    GM_xmlhttpRequest({
      method:'GET', url:location.href, responseType:'text', timeout:15000,
      onload:function(r){
        var m=(r.responseText||'').match(/shopId=([A-Za-z0-9]+)/);
        if(m && m[1]){
          var _mode='menu';
          try{ _mode=localStorage.getItem('cmz_mode')||'menu'; }catch(e){}
          banner('✅ 已识别店铺，正在进入并自动抓取（模式：'+_mode+'）…');
          location.replace(buildShopUrl(m[1], _mode));
        }else{
          banner('❌ 未能从口令中解析出店铺。请复制地址栏链接（含 shopId=…）粘贴到网站，或用手机打开口令后把 h5.ele.me 链接粘贴回来。','#FEF2F2');
        }
      },
      onerror:function(){ banner('❌ 口令解析失败（网络原因）。请复制地址栏链接（含 shopId=…）粘贴到网站。','#FEF2F2'); }
    });
    return;
  }

  /* m.tb.cn 跳转下载页：提取 shopId 并跳转店铺 */
  if(host.indexOf('market.m.taobao.com') >= 0){
    var sid = getParam('shopId');
    if(sid){
      var m = 'menu';
      try{ m = localStorage.getItem('cmz_mode') || 'menu'; }catch(e){}
      banner('✅ 检测到店铺链接，正在进入店铺并自动抓取（模式：'+m+'）…');
      location.replace(buildShopUrl(sid, m));
    }
    return;
  }

  /* 店铺页（ele.me）带 cmz_scrape 参数：自动抓取 */
  if(host.indexOf('ele.me') >= 0 && mode){
    (async function(){
      try{ localStorage.setItem('cmz_mode', mode); }catch(e){}

      /* 风控检测：若已跳转到安全验证页，明确提示登录 */
      if(/punish|waimai-guide|x5secdata/i.test(location.href)){
        banner('⚠️ 触发平台安全验证。请先完成登录：打开登录页后用手机 App 扫码，再重新点「开始抓取」。若仍出现，请稍等几分钟再试。', '#FEF3C7');
        return;
      }

      banner('⏳ 正在识别店铺菜单，请稍候…（自动抓取中，无需操作）');
      // 等待菜单出现
      var waited=0;
      while(waited<30000){
        if(document.querySelectorAll('.menuItem, .food-item--wrap').length>0) break;
        if(/punish|waimai-guide|x5secdata/i.test(location.href)){
          banner('⚠️ 触发平台安全验证。请先完成登录（手机 App 扫码）后重新抓取。', '#FEF3C7');
          return;
        }
        await sleep(500); waited+=500;
      }
      await loadAll();
      if(/punish|waimai-guide|x5secdata/i.test(location.href)){
        banner('⚠️ 抓取过程中触发平台安全验证，已自动停止。请先完成登录后重新点「开始抓取」。', '#FEF3C7');
        return;
      }
      var items = parseMenu();
      if(!items.length){
        banner('❌ 未识别到菜单，请确认当前页面是店铺「点餐」页。若是口令链接，请先让本插件自动跳转后再试。', '#FEF2F2');
        return;
      }
      banner('✅ 识别到 ' + items.length + ' 个菜品，正在生成文件…');
      var shopName = '';
      try{
        var t=(document.body.innerText||'');
        var m2=t.match(/([^\n]{2,20}?\([^)]*店\))/) || t.match(/赣味[^\n]{0,15}|[^\n]{2,15}店/);
        if(m2) shopName = esc(m2[1]).slice(0,30);
      }catch(e){}
      var fnBase = (shopName||'店铺').replace(/\s+/g,'') + '_' + ts();

      /* 菜单 Excel */
      if(mode==='menu' || mode==='all'){
        var rows=[['分组','菜品名称','现价(元)','券后/预估到手(元)','月售','说明','有无图片']];
        items.forEach(function(x){
          rows.push([x.cate, x.name, x.price||'', x.useCoupon||'', x.sales, x.desc, x.img?'有':'']);
        });
        var xlsx=buildMenuXlsx(rows);
        downloadBlob(new Blob([xlsx],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}), '抓取菜单_'+fnBase+'.xlsx');
      }

      /* 菜品图 zip */
      if(mode==='img' || mode==='all'){
        var imgs = items.filter(function(x){ return x.img; });
        banner('✅ 菜单已生成。正在下载 ' + imgs.length + ' 张菜品图（zip 打包中）…');
        var files=[];
        for(var i=0;i<imgs.length;i++){
          try{
            var it=imgs[i];
            // 原图：去掉缩略图处理参数
            var u=it.img.split('?')[0];
            var blob;
            try{ blob = await gmFetch(u, 'blob'); }
            catch(err){ blob = await gmFetch(it.img, 'blob'); }
            var ab = await blob.arrayBuffer();
            var ext='jpg'; var mm=u.match(/\.(jpe?g|png|webp|gif)$/i); if(mm) ext=mm[1].toLowerCase();
            files.push({name: esc(it.cate||'未分组')+'_'+esc(it.name).slice(0,30)+'.'+ext, data:new Uint8Array(ab)});
          }catch(e){}
        }
        if(files.length){
          var z=buildZip(files);
          downloadBlob(new Blob([z],{type:'application/zip'}), '菜品图_'+fnBase+'.zip');
          banner('✅ 全部完成：' + items.length + ' 个菜品'+(mode==='all'?'，Excel 已下载':'')+'，图片 '+files.length+' 张已打包下载。可关闭本页。', '#F0FDF4');
        } else {
          banner('⚠️ 菜品图下载失败（可能被平台防盗链拦截），Excel 已生成可正常使用。', '#FEF3C7');
        }
      } else {
        banner('✅ 抓取完成：' + items.length + ' 个菜品，Excel 已自动下载。可关闭本页。', '#F0FDF4');
      }
    })();
  }
})();