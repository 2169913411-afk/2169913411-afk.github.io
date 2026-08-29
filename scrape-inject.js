/* 餐谋长·运营助手 - 商家调研零安装抓取脚本 v3.3
   v3.3 更新：
   - 新增美团外卖支持（自动识别美团页面，使用美团选择器和抓取逻辑）
   - 美团抓取支持菜单、图片、菜单+图片三种模式
   v3.2 更新：
   - 抓取完成后不再自动下载，改为顶部下载面板（用户点击下载，可自选保存位置 showSaveFilePicker）
   - 菜品图文件名去掉分组前缀，只保留菜品名
   v3.1：修复菜品图（CSS background-image 解析 + 取高清原图）
   由「餐谋长抓取」书签在店铺页动态加载执行：
   1) 读取 URL 参数 cmz_scrape=menu|img|all（网站在打开店铺时已自动写入）
   2) 自动识别平台（淘宝闪购/美团外卖），使用对应选择器
   3) 依次点击左侧全部分类tab + 滚动到底，加载全部菜品
   4) 菜单 -> CSV(Excel可打开)；菜品图 -> zip 打包下载（高清原图）
   仅抓取用户端公开展示的店铺菜单数据，用于商家调研分析。 */
(function(){
  'use strict';
  var BRAND='餐谋长·抓取';

  /* ---------- 平台检测 ---------- */
  function isMeituanPage(){
    return /waimai\.meituan\.com|meituan\.com\/waimai/i.test(location.href);
  }
  function isElemePage(){
    return /h5\.ele\.me|ele\.me/i.test(location.href);
  }
  function banner(msg, color){
    var id='cmz-banner';
    var b=document.getElementById(id);
    if(!b){
      b=document.createElement('div');
      b.id=id;
      b.style.cssText='position:fixed;top:0;left:0;right:0;z-index:2147483647;padding:10px 16px;font-size:14px;font-weight:600;font-family:-apple-system,"PingFang SC",sans-serif;text-align:center;box-shadow:0 2px 8px rgba(0,0,0,.15);';
      document.documentElement.appendChild(b);
    }
    b.style.background=color||'#EFF6FF';
    b.style.color='#1D4ED8';
    b.style.borderBottom='2px solid #2563EB';
    b.innerHTML=msg;
  }
  function sleep(ms){ return new Promise(function(r){ setTimeout(r,ms); }); }
  function numText(s){ return String(s||'').replace(/[^\d.]/g,''); }
  function txt(item,sel){
    var el=item.querySelector(sel);
    return el?(el.innerText||'').replace(/\s+/g,'').trim():'';
  }
  function escCSV(s){
    s=String(s==null?'':s).replace(/\r?\n/g,' ');
    if(/[",\n]/.test(s)) s='"'+s.replace(/"/g,'""')+'"';
    return s;
  }
  function ts(){ var d=new Date(); function p2(x){return ('0'+x).slice(-2);} return d.getFullYear()+p2(d.getMonth()+1)+p2(d.getDate())+'_'+p2(d.getHours())+p2(d.getMinutes()); }
  function downloadBlob(blob, name){
    var url=URL.createObjectURL(blob);
    var a=document.createElement('a');
    a.href=url; a.download=name;
    document.body.appendChild(a); a.click();
    setTimeout(function(){ document.body.removeChild(a); URL.revokeObjectURL(url); },400);
  }
  /* 保存文件：优先系统「另存为」自选保存位置；不可用则退回浏览器默认下载 */
  async function saveAs(blob, suggestedName){
    if(window.showSaveFilePicker){
      try{
        var handle=await window.showSaveFilePicker({suggestedName:suggestedName});
        var w=await handle.createWritable();
        await w.write(blob);
        await w.close();
        return true;
      }catch(e){
        if(e&&e.name==='AbortError') return false; // 用户取消
      }
    }
    downloadBlob(blob, suggestedName);
    return true;
  }
  window.__cmzSave=async function(key){
    var r=window.__cmzResult; if(!r) return;
    if(key==='menu'&&r.menuText){
      await saveAs(new Blob(['\ufeff'+r.menuText],{type:'text/csv;charset=utf-8'}), '抓取菜单_'+r.fnBase+'.csv');
    }else if(key==='zip'&&r.zip){
      await saveAs(r.zip, '菜品图_'+r.fnBase+'.zip');
    }
  };
  function isPunish(){ return /punish|waimai-guide|x5secdata/i.test(location.href); }
  function getParam(name){
    var m=location.href.match(new RegExp('[?&#]'+name+'=([^&#]*)'));
    return m?decodeURIComponent(m[1]):'';
  }
  function loadJS(src){
    return new Promise(function(res,rej){
      if(document.querySelector('script[data-cmzjs="'+src+'"]')){ res(); return; }
      var s=document.createElement('script');
      s.setAttribute('data-cmzjs',src);
      s.src=src;
      s.onload=function(){ res(); };
      s.onerror=function(){ rej(new Error('加载失败:'+src)); };
      document.head.appendChild(s);
    });
  }
  /* 取高清原图：去掉 oss 处理参数 */
  function cleanImgUrl(u){
    if(!u) return '';
    u=u.replace(/^["']|["']$/g,'');
    var idx=u.indexOf('?x-oss-process');
    if(idx>-1) u=u.slice(0,idx);
    return u;
  }
  /* 从菜单条目提取菜品图：饿了么 H5 菜品图是 CSS background-image */
  function getItemImg(item){
    var sel=['.menuItem--image-img','.menuItem--image','[class*="image-img"]','[class*="menuItem--image"]'];
    for(var k=0;k<sel.length;k++){
      var els=item.querySelectorAll(sel[k]);
      for(var i=0;i<els.length;i++){
        var cs=window.getComputedStyle(els[i]);
        var m=cs.backgroundImage&&cs.backgroundImage.match(/url\(["']?([^"')]+)["']?\)/);
        if(m&&m[1]&&/elemecdn|alicdn|cube\./.test(m[1])) return cleanImgUrl(m[1]);
      }
    }
    var img=item.querySelector('img');
    if(img){
      var s=img.getAttribute('src')||img.getAttribute('data-src')||'';
      if(s) return cleanImgUrl(s);
    }
    var all=item.querySelectorAll('*');
    for(var j=0;j<all.length;j++){
      var cs2=window.getComputedStyle(all[j]);
      var m2=cs2.backgroundImage&&cs2.backgroundImage.match(/url\(["']?([^"')]+)["']?\)/);
      if(m2&&m2[1]&&/elemecdn|alicdn|cube\./.test(m2[1])) return cleanImgUrl(m2[1]);
    }
    return '';
  }
  /* 解析菜单条目（含图片地址） */
  function parseMenu(){
    var items=[], seen={};
    document.querySelectorAll('.menuItem').forEach(function(item){
      var cate=item.getAttribute('data-cate-name')||'';
      var title=txt(item,'.menuItem--info-title')||txt(item,'.menuItem--info-title--warp');
      var price=txt(item,'.menuItem--info-price-text')||numText(txt(item,'.menuItem--info-price'));
      var use=txt(item,'.menuItem--info-useCouponPrice');
      var sales=txt(item,'.menuItem--info-sales');
      var desc=txt(item,'.menuItem--info-description');
      var img=getItemImg(item);
      if(!title) return;
      var key=title+'|'+cate;
      if(seen[key]) return; seen[key]=1;
      items.push({cate:cate,name:title,price:price,use:use,sales:sales,desc:desc,img:img});
    });
    return items;
  }
  /* 依次点击左侧全部分类tab */
  async function clickAllTabs(){
    var tabs=document.querySelectorAll('.sideList--item');
    if(!tabs.length) return true;
    for(var i=0;i<tabs.length;i++){
      if(isPunish()) return false;
      try{ tabs[i].click(); }catch(e){}
      await sleep(900+Math.random()*500);
    }
    return true;
  }
  /* 滚动到最底部直到菜品数量稳定 */
  async function loadAll(){
    var sc=document.querySelector('.mor-comp-page-content');
    if(!sc){
      var all=document.querySelectorAll('*');
      for(var i=0;i<all.length;i++){
        var el=all[i];
        if(el.scrollHeight>el.clientHeight+200 && el.clientHeight>100){ sc=el; break; }
      }
    }
    if(!sc) return true;
    var prev=-1, stable=0;
    for(var i=0;i<50 && stable<4; i++){
      sc.scrollTop=sc.scrollHeight;
      await sleep(700+Math.random()*500);
      if(isPunish()) return false;
      var n=document.querySelectorAll('.menuItem').length;
      if(n===prev){ stable++; } else { stable=0; prev=n; }
    }
    sc.scrollTop=0;
    return true;
  }
  /* 下载图片并打包 zip（返回 {ok, zip}，不自动下载） */
  async function buildImgsZip(items){
    var imgs=items.filter(function(x){ return x.img; });
    if(!imgs.length) return {ok:0, zip:null};
    banner('⏳ '+BRAND+'：正在下载 '+imgs.length+' 张菜品图并打包…');
    try{ await loadJS('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'); }
    catch(e){ return {ok:-1, zip:null}; }
    if(typeof JSZip==='undefined') return {ok:-1, zip:null};
    var zip=new JSZip(), ok=0;
    for(var i=0;i<imgs.length;i++){
      if(isPunish()) return {ok:-2, zip:null};
      try{
        var it=imgs[i];
        var u=(it.img.split('?')[0])||it.img;
        var r=await fetch(u, {mode:'cors'});
        if(!r.ok) throw new Error('http'+r.status);
        var blob=await r.blob();
        var ext='jpg'; var mm=u.match(/\.(jpe?g|png|webp|gif)$/i); if(mm) ext=mm[1].toLowerCase();
        /* v3.2：文件名只保留菜品名，不带分组前缀 */
        var name=(it.name||('img'+i)).replace(/[\\\/:*?"<>|\r\n]/g,' ').slice(0,40)+'.'+ext;
        zip.file(name, blob);
        ok++;
      }catch(e){}
      if(i%5===0) banner('⏳ '+BRAND+'：图片下载中 '+i+'/'+imgs.length+'…');
    }
    var z=ok?await zip.generateAsync({type:'blob'}):null;
    return {ok:ok, zip:z};
  }
  /* 抓取完成：显示下载面板（用户点按钮时自选保存位置） */
  function showDownloadPanel(r){
    window.__cmzResult=r;
    var parts=[];
    parts.push('✅ '+BRAND+'：识别到 <b>'+r.itemsCount+'</b> 个菜品（覆盖 '+r.cateCount+' 个分类）');
    if(r.menuText) parts.push('<button style="margin:0 6px;padding:7px 16px;border:none;border-radius:8px;background:#165DFF;color:#fff;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 6px rgba(22,93,255,.3);" onclick="window.__cmzSave(\'menu\')">下载菜单（Excel）</button>');
    if(r.zip) parts.push('<button style="margin:0 6px;padding:7px 16px;border:none;border-radius:8px;background:#FF7D00;color:#fff;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 6px rgba(255,125,0,.3);" onclick="window.__cmzSave(\'zip\')">下载菜品图包（zip·'+r.imgCount+'张）</button>');
    parts.push('<span style="font-size:11px;opacity:.7;">点按钮下载可自选保存位置</span>');
    banner(parts.join(' '),'#F0FDF4');
  }
  (async function(){
    /* ---------- 美团外卖：动态加载美团专用抓取脚本 ---------- */
    if(isMeituanPage()){
      banner('⏳ '+BRAND+'：识别到美团外卖店铺，正在加载美团抓取脚本…');
      try{
        await loadJS('https://2169913411-afk.github.io/meituan-scraper.js?v=3.0');
        return; // 美团脚本会自行处理抓取流程
      }catch(e){
        banner('❌ '+BRAND+'：美团抓取脚本加载失败，请检查网络后重试。','#FEF2F2');
        return;
      }
    }

    if(isPunish()){
      banner('⚠️ '+BRAND+'：当前是平台安全验证页。请先用手机「饿了么」App 扫码登录后，再重新打开店铺点书签。','#FEF3C7');
      return;
    }
    var mode=getParam('cmz_scrape')||'menu';
    if(mode!=='menu'&&mode!=='img'&&mode!=='all') mode='menu';
    var mname=mode==='menu'?'菜单':(mode==='img'?'菜品图':'菜单+菜品图');
    banner('⏳ '+BRAND+'：正在展开全部菜单（遍历分类+滚动加载），抓取内容：'+mname+'…');
    var waited=0;
    while(waited<30000){
      if(document.querySelectorAll('.food-item--wrap,.menuItem,.sideList--item').length>0) break;
      if(isPunish()){
        banner('⚠️ '+BRAND+'：触发平台安全验证，请登录后重试。','#FEF3C7');
        return;
      }
      await sleep(500); waited+=500;
    }
    var ok=await clickAllTabs();
    if(!ok){ banner('⚠️ '+BRAND+'：加载中触发安全验证，已停止。','#FEF3C7'); return; }
    ok=await loadAll();
    if(!ok){ banner('⚠️ '+BRAND+'：加载中触发安全验证，已停止。','#FEF3C7'); return; }
    if(isPunish()){
      banner('⚠️ '+BRAND+'：抓取中触发平台安全验证，已自动停止，请登录后重试。','#FEF3C7');
      return;
    }
    var items=parseMenu();
    if(!items.length){
      banner('❌ '+BRAND+'：未识别到菜单。请确认当前页面是店铺「点餐」页后，再点一次书签。','#FEF2F2');
      return;
    }
    var shop='';
    var t=document.body.innerText||'';
    var m=t.match(/([^\n]{2,20}?\([^)]*店\))/)||t.match(/[^\n]{2,15}店/);
    if(m) shop=String(m[1]||'').replace(/[\\\/:*?"<>|\n\r]/g,' ').trim().slice(0,30);
    var fnBase=(shop||'店铺').replace(/\s+/g,'')+'_'+ts();
    var cateSet={}; items.forEach(function(x){ if(x.cate) cateSet[x.cate]=1; });
    var result={menuText:null, zip:null, imgCount:0, itemsCount:items.length, cateCount:Object.keys(cateSet).length, fnBase:fnBase, shop:shop};
    /* 菜单 */
    if(mode==='menu'||mode==='all'){
      var lines=['分组,菜品名称,现价(元),券后/预估到手(元),月售,说明,有无图片'];
      items.forEach(function(x){
        lines.push([escCSV(x.cate),escCSV(x.name),escCSV(x.price),escCSV(x.use),escCSV(x.sales),escCSV(x.desc),x.img?'有':''].join(','));
      });
      result.menuText=lines.join('\n');
    }
    /* 图片 */
    if(mode==='img'||mode==='all'){
      var g=await buildImgsZip(items);
      if(g.ok===-2){ banner('⚠️ '+BRAND+'：图片下载中触发安全验证，菜单部分可点上方按钮下载。','#FEF3C7'); }
      result.zip=g.ok>0?g.zip:null;
      result.imgCount=g.ok>0?g.ok:0;
    }
    showDownloadPanel(result);
  })();
})();
