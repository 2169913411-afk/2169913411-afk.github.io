/* 餐谋长·运营助手 - 商家调研零安装抓取脚本 v3（支持 URL 参数自动模式：菜单/图片/全部）
   由「餐谋长抓取」书签在店铺页动态加载执行：
   1) 读取 URL 参数 cmz_scrape=menu|img|all（网站在打开店铺时已自动写入）
   2) 依次点击左侧全部分类tab + 滚动到底，加载全部菜品
   3) 菜单 -> CSV(Excel可打开)；菜品图 -> zip 打包下载
   仅抓取用户端公开展示的店铺菜单数据，用于商家调研分析。 */
(function(){
  'use strict';
  var BRAND='餐谋长·抓取';
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
  function downloadCSV(text, name){
    downloadBlob(new Blob(['\ufeff'+text],{type:'text/csv;charset=utf-8'}), name);
  }
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
      var imgEl=item.querySelector('img');
      var img=imgEl?(imgEl.getAttribute('src')||imgEl.getAttribute('data-src')||''):'';
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
  /* 下载图片 zip */
  async function grabImgs(items, fnBase){
    var imgs=items.filter(function(x){ return x.img; });
    if(!imgs.length) return 0;
    banner('⏳ '+BRAND+'：正在下载 '+imgs.length+' 张菜品图并打包…');
    try{ await loadJS('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js'); }
    catch(e){ banner('❌ '+BRAND+'：压缩组件加载失败，图片未打包。','#FEF2F2'); return -1; }
    if(typeof JSZip==='undefined'){ banner('❌ '+BRAND+'：压缩组件不可用。','#FEF2F2'); return -1; }
    var zip=new JSZip(), ok=0, fail=0;
    for(var i=0;i<imgs.length;i++){
      if(isPunish()) return -2;
      try{
        var it=imgs[i];
        var u=(it.img.split('?')[0])||it.img;
        var r=await fetch(u, {mode:'cors'});
        if(!r.ok) throw new Error('http'+r.status);
        var blob=await r.blob();
        var ext='jpg'; var mm=u.match(/\.(jpe?g|png|webp|gif)$/i); if(mm) ext=mm[1].toLowerCase();
        var name=(it.cate||'未分组')+'_'+(it.name||('img'+i)).replace(/[\\\/:*?"<>|\r\n]/g,' ').slice(0,30)+'.'+ext;
        zip.file(name, blob);
        ok++;
      }catch(e){ fail++; }
      if(i%5===0) banner('⏳ '+BRAND+'：图片下载中 '+i+'/'+imgs.length+'…');
    }
    if(ok){
      var z=await zip.generateAsync({type:'blob'});
      downloadBlob(z, '菜品图_'+fnBase+'.zip');
    }
    return ok;
  }
  (async function(){
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
    /* 菜单 */
    if(mode==='menu'||mode==='all'){
      var lines=['分组,菜品名称,现价(元),券后/预估到手(元),月售,说明,有无图片'];
      items.forEach(function(x){
        lines.push([escCSV(x.cate),escCSV(x.name),escCSV(x.price),escCSV(x.use),escCSV(x.sales),escCSV(x.desc),x.img?'有':''].join(','));
      });
      downloadCSV(lines.join('\n'),'抓取菜单_'+fnBase+'.csv');
    }
    /* 图片 */
    var imgsResult=0;
    if(mode==='img'||mode==='all'){
      imgsResult=await grabImgs(items, fnBase);
    }
    if(imgsResult===-2){ banner('⚠️ '+BRAND+'：图片下载中触发安全验证，菜单部分已完成。','#FEF3C7'); return; }
    var doneImg=imgsResult>0?('，图片 '+imgsResult+' 张已打包'):(imgsResult===-1?'，图片打包失败(网络/组件)':'');
    var cateSet={}; items.forEach(function(x){ if(x.cate) cateSet[x.cate]=1; });
    banner('✅ '+BRAND+'：识别到 '+items.length+' 个菜品（覆盖 '+Object.keys(cateSet).length+' 个分类）'+doneImg+'。文件已自动下载，可用 Excel 打开。可关闭本页。','#F0FDF4');
  })();
})();
