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

  function isJDMerchantPage(){
    return /^store\.jd\.com$/i.test(location.hostname) && /\/product\/shop\/index/.test(location.pathname);
  }
  function jdText(el){ return el ? String(el.innerText || el.textContent || '').replace(/\u00a0/g,' ').trim() : ''; }
  function jdNorm(s){ return String(s||'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim(); }
  function jdVisible(el){
    if(!el) return false;
    var r=el.getBoundingClientRect(), st=window.getComputedStyle(el);
    return r.width>0 && r.height>0 && st.display!=='none' && st.visibility!=='hidden';
  }
  function jdSkipCategory(name){ return /^(全部商品|全部商品分类)$/.test(jdNorm(name)); }
  function jdAllCategoryNode(){
    var a=Array.from(document.querySelectorAll('body *')).filter(function(el){
      if(!jdVisible(el)) return false;
      var t=jdNorm(jdText(el)), r=el.getBoundingClientRect();
      return jdSkipCategory(t) && r.left < window.innerWidth*0.48 && r.top > 50;
    });
    a.sort(function(x,y){
      var rx=x.getBoundingClientRect(), ry=y.getBoundingClientRect();
      return (rx.width*rx.height)-(ry.width*ry.height);
    });
    return a[0]||null;
  }
  function jdCategories(anchor){
    if(!anchor) return [];
    var ar=anchor.getBoundingClientRect(), out=[], seen={};
    Array.from(document.querySelectorAll('body *')).forEach(function(el){
      if(!jdVisible(el)) return;
      var r=el.getBoundingClientRect(), label=jdNorm(jdText(el));
      if(!label || label.length>32 || r.left<ar.left-35 || r.left>ar.left+150 || r.top<ar.top-8) return;
      if(jdSkipCategory(label)) return;
      if(/^(商品管理|全部分类|店铺商品管理)$/.test(label)) return;
      if(!seen[label]){ seen[label]=1; out.push(label); }
    });
    return out;
  }
  function jdFindCategoryNode(name, anchor){
    if(!anchor) return null;
    var ar=anchor.getBoundingClientRect(), a=Array.from(document.querySelectorAll('body *')).filter(function(el){
      if(!jdVisible(el)) return false;
      var r=el.getBoundingClientRect();
      return jdNorm(jdText(el))===name && r.left>=ar.left-35 && r.left<=ar.left+150 && r.top>=ar.top-8;
    });
    a.sort(function(x,y){
      var rx=x.getBoundingClientRect(), ry=y.getBoundingClientRect();
      return (rx.width*rx.height)-(ry.width*ry.height);
    });
    return a[0]||null;
  }
  function jdSkuIds(s){
    var out=[], re=/(?:SKU\s*编码|SKU编码)\s*[:：]?\s*([A-Za-z0-9_-]+)/ig, m;
    while((m=re.exec(String(s||'')))!==null) if(out.indexOf(m[1])<0) out.push(m[1]);
    return out;
  }
  function jdProductDetails(){
    var els=Array.from(document.querySelectorAll('tr,[role="row"],[class*="row"],[class*="Row"],td,div'));
    var candidates=els.filter(function(el){
      var t=jdText(el);
      return jdVisible(el) && /SKU\s*编码|SKU编码/i.test(t) && /SPU\s*编码|SPU编码/i.test(t) && jdSkuIds(t).length===1;
    });
    candidates.sort(function(a,b){
      var x=a.getBoundingClientRect(), y=b.getBoundingClientRect();
      return (x.width*x.height)-(y.width*y.height);
    });
    var out=[], seen={};
    candidates.forEach(function(detail){
      var raw=jdText(detail), ids=jdSkuIds(raw), id=ids[0];
      if(!id || seen[id]) return;
      seen[id]=1;
      var row=detail, p=detail.parentElement;
      for(var i=0;p && i<7 && p!==document.body;i++,p=p.parentElement){
        var pt=jdText(p);
        if(jdSkuIds(pt).length>1) break;
        if(/(?:门店价格|实际价格|售价|销售价)\s*[:：]?\s*[¥￥]?\s*\d/i.test(pt) || p.matches('tr,[role="row"],[class*="row"],[class*="Row"]')){ row=p; break; }
        row=p;
      }
      var detailRaw=jdText(detail).replace(/(SKU\s*编码|SPU\s*编码|商家\s*sku\s*编码|门店价格|实际价格|售价|销售价)\s*[:：]/ig,'\n$1：');
      var lines=detailRaw.split(/\n+/).map(jdNorm).filter(Boolean);
      var skuLine=lines.findIndex(function(x){ return /SKU\s*编码|SKU编码/i.test(x); });
      var titleMatch=detailRaw.match(/(?:商品名称|商品名)\s*[:：]\s*([^\n]+)/);
      var before=skuLine>=0?lines.slice(0,skuLine):lines;
      before=before.filter(function(x){ return !/^(商品信息|价格信息|SKU编码|SPU编码|商家sku编码|门店ID|店铺ID)/i.test(x); });
      var name=titleMatch?jdNorm(titleMatch[1]):before.sort(function(x,y){return y.length-x.length;})[0]||'';
      name=name.replace(/^(商品名称|商品名)\s*[:：]\s*/,'').trim();
      var rowRaw=jdText(row), pm=rowRaw.match(/(?:门店价格|实际价格|售价|销售价|商品价格)\s*[:：]?\s*[¥￥]?\s*(\d+(?:\.\d{1,2})?)/i);
      var img='';
      var im=row.querySelector && row.querySelector('img');
      if(im) img=im.getAttribute('src')||im.getAttribute('data-src')||'';
      if(!img && detail.querySelector){ im=detail.querySelector('img'); if(im) img=im.getAttribute('src')||im.getAttribute('data-src')||''; }
      out.push({row:row,sku:id,name:name,price:pm?pm[1]:'',desc:'',img:img});
    });
    return out;
  }
  function jdScrollElement(){
    var root=document.scrollingElement;
    if(root && root.scrollHeight>root.clientHeight+120) return root;
    var a=Array.from(document.querySelectorAll('div,[class*="scroll"],[class*="Scroll"]')).filter(function(el){
      var st=window.getComputedStyle(el), r=el.getBoundingClientRect();
      return jdVisible(el) && /(auto|scroll)/.test(st.overflowY) && el.scrollHeight>el.clientHeight+120 && r.width>window.innerWidth*.35 && r.height>180;
    });
    a.sort(function(x,y){
      var rx=x.getBoundingClientRect(), ry=y.getBoundingClientRect();
      return (ry.width*ry.height)-(rx.width*rx.height);
    });
    return a[0]||root;
  }
  function jdNextPageButton(){
    var direct=document.querySelector('.ant-pagination-next:not(.ant-pagination-disabled) button,.ant-pagination-next:not(.ant-pagination-disabled)');
    if(direct && jdVisible(direct) && !direct.disabled) return direct;
    return Array.from(document.querySelectorAll('button,a,[role="button"]')).find(function(el){
      var t=jdNorm(jdText(el)), r=el.getBoundingClientRect();
      return jdVisible(el) && r.top>0 && /^(下一页|下一页›|›|>)+$/.test(t) && !el.disabled && !/disabled/i.test(el.className||'');
    })||null;
  }
  async function jdWaitChange(before, ms){
    var elapsed=0;
    while(elapsed<ms){
      if(isPunish() || /验证码|安全验证|操作过于频繁/.test(document.body.innerText||'')) return false;
      await sleep(250); elapsed+=250;
      var now=jdProductDetails().map(function(x){return x.sku+':'+x.name;}).join('|');
      if(now && now!==before) return true;
    }
    return true;
  }
  async function jdCollectCategory(category){
    var found={}, sc=jdScrollElement(), original=sc?sc.scrollTop:0, page=0;
    for(page=0;page<60;page++){
      if(isPunish() || /验证码|安全验证|操作过于频繁/.test(document.body.innerText||'')) break;
      var stable=0, previousCount=-1, steps=0;
      while(sc && steps<60 && stable<3){
        var rows=jdProductDetails();
        rows.forEach(function(x){
          if(!x.name) return;
          var key=category+'|'+x.sku;
          if(!found[key]) found[key]={cate:category,name:x.name,price:x.price,sales:'',use:'',desc:x.desc,img:x.img,sku:x.sku};
        });
        var count=Object.keys(found).length;
        if(count===previousCount) stable++; else stable=0;
        previousCount=count;
        if(sc.scrollTop+sc.clientHeight>=sc.scrollHeight-8) break;
        var old=sc.scrollTop;
        sc.scrollTop=Math.min(sc.scrollHeight,old+Math.max(320,Math.floor(sc.clientHeight*.8)));
        await sleep(350);
        if(sc.scrollTop===old) break;
        steps++;
      }
      if(!sc) jdProductDetails().forEach(function(x){
        if(x.name){ var key=category+'|'+x.sku; if(!found[key]) found[key]={cate:category,name:x.name,price:x.price,sales:'',use:'',desc:x.desc,img:x.img,sku:x.sku}; }
      });
      if(sc) sc.scrollTop=original;
      var next=jdNextPageButton();
      if(!next || page>=59) break;
      var before=jdProductDetails().map(function(x){return x.sku+':'+x.name;}).join('|');
      next.click();
      await jdWaitChange(before,6000);
      if(isPunish() || /验证码|安全验证|操作过于频繁/.test(document.body.innerText||'')) break;
    }
    if(sc) sc.scrollTop=original;
    return Object.keys(found).map(function(k){return found[k];});
  }
  async function runJDMerchant(){
    if(isPunish() || /验证码|安全验证|操作过于频繁/.test(document.body.innerText||'')){
      banner('⚠️ 检测到京东安全验证或频繁操作提示，已停止采集，请按平台页面提示处理后再试。','#FEF3C7'); return;
    }
    var all=jdAllCategoryNode(), categories=jdCategories(all);
    categories=categories.filter(function(x){return !jdSkipCategory(x);});
    if(!all || !categories.length){
      banner('❌ 未识别到京东门店商品分类。请停留在「门店商品管理」页面后重新点书签。','#FEF2F2'); return;
    }
    var items=[];
    banner('⏳ 京东门店商品采集中：已排除「全部商品」，共识别 '+categories.length+' 个其他分类…');
    for(var i=0;i<categories.length;i++){
      if(isPunish() || /验证码|安全验证|操作过于频繁/.test(document.body.innerText||'')){
        banner('⚠️ 检测到京东安全验证或频繁操作提示，已停止。当前已采集 '+items.length+' 件，请处理页面提示后重试。','#FEF3C7'); return;
      }
      var node=jdFindCategoryNode(categories[i],all);
      if(!node) continue;
      var before=jdProductDetails().map(function(x){return x.sku+':'+x.name;}).join('|');
      node.click();
      await jdWaitChange(before,6000);
      var part=await jdCollectCategory(categories[i]);
      items=items.concat(part);
      banner('⏳ 京东采集中：'+(i+1)+'/'+categories.length+' 个分类，已收集 '+items.length+' 件…');
    }
    if(!items.length){
      banner('❌ 页面未识别到带 SKU/SPU 的商品行，未生成菜单。请检查商品列表是否加载完成。','#FEF2F2'); return;
    }
    var seenCate={};
    items.forEach(function(x){ if(x.cate) seenCate[x.cate]=1; });
    var lines=[['分组','商品名称','SKU编码','售价(元)','商品描述','图片链接'].map(escCSV).join(',')];
    items.forEach(function(x){ lines.push([x.cate,x.name,x.sku,x.price,x.desc,x.img].map(escCSV).join(',')); });
    var t=document.body.innerText||'', m=t.match(/[^\n]{2,30}(?:店|门店)/), shop=m?jdNorm(m[0]):'京东门店';
    var fnBase=shop.replace(/[\\/:*?"<>|]/g,' ').slice(0,40)+'_'+ts();
    var result={menuText:lines.join('\n'),zip:null,imgCount:0,itemsCount:items.length,cateCount:Object.keys(seenCate).length,fnBase:fnBase,shop:shop,items:items};
    var mode=getParam('cmz_scrape')||'menu';
    if(mode==='img'||mode==='all'){
      var zipped=await buildImgsZip(items);
      result.zip=zipped.ok>0?zipped.zip:null;
      result.imgCount=zipped.ok>0?zipped.ok:0;
    }
    showDownloadPanel(result);
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
        if(m&&m[1]&&/elemecdn|alicdn|cube\.|360buyimg|jdimg/.test(m[1])) return cleanImgUrl(m[1]);
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
      if(m2&&m2[1]&&/elemecdn|alicdn|cube\.|360buyimg|jdimg/.test(m2[1])) return cleanImgUrl(m2[1]);
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
  /* 抓取完成：显示下载面板（用户点按钮时自选保存位置）+ 存储数据到localStorage + 返回网站按钮 */
  function showDownloadPanel(r){
    window.__cmzResult=r;
    
    /* v3.5：把抓取的菜单数据存储到window.name（跨域传递，跳转后保留） */
    try{
      var shopData = {
        platform: isJDMerchantPage() ? '京东秒送' : (isMeituanPage() ? '美团外卖' : '淘宝闪购'),
        shopName: r.shop || '',
        shopUrl: location.href,
        scrapeTime: new Date().toISOString(),
        categories: [],
        products: []
      };
      
      /* 从items中提取分类和商品 */
      var cateMap = {};
      if(r.items && r.items.length){
        r.items.forEach(function(item){
          if(item.cate && !cateMap[item.cate]){
            cateMap[item.cate] = {name: item.cate, products: []};
            shopData.categories.push(cateMap[item.cate]);
          }
          var product = {
            name: item.name || '',
            price: item.price ? parseFloat(item.price) : null,
            originalPrice: item.use ? parseFloat(item.use) : null,
            monthSales: item.sales || '',
            description: item.desc || '',
            image: item.img || '',
            sku: item.sku || '',
            category: item.cate || ''
          };
          shopData.products.push(product);
          if(cateMap[item.cate]){
            cateMap[item.cate].products.push(product);
          }
        });
      }
      
      /* 存储到window.name（跨域传递，跳转后保留） */
      var dataStr = JSON.stringify(shopData);
      window.name = 'CMZ_SCRAPED_DATA:' + dataStr;
      console.log('餐谋长抓取：数据已存储到window.name，共', shopData.products.length, '个商品，大小', dataStr.length, '字节');
    }catch(e){
      console.error('存储数据失败:', e);
    }
    
    var parts=[];
    parts.push('✅ '+BRAND+'：识别到 <b>'+r.itemsCount+'</b> 个菜品（覆盖 '+r.cateCount+' 个分类）');
    
    /* v3.4：添加返回餐谋长网站按钮 */
    parts.push('<button style="margin:0 6px;padding:7px 16px;border:none;border-radius:8px;background:#16A34A;color:#fff;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 6px rgba(22,163,74,.3);" onclick="window.__cmzReturnToSite()">← 返回餐谋长网站（自动导入数据）</button>');
    
    if(r.menuText) parts.push('<button style="margin:0 6px;padding:7px 16px;border:none;border-radius:8px;background:#165DFF;color:#fff;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 6px rgba(22,93,255,.3);" onclick="window.__cmzSave(\'menu\')">下载菜单（Excel）</button>');
    if(r.zip) parts.push('<button style="margin:0 6px;padding:7px 16px;border:none;border-radius:8px;background:#FF7D00;color:#fff;font-size:13px;font-weight:600;cursor:pointer;box-shadow:0 2px 6px rgba(255,125,0,.3);" onclick="window.__cmzSave(\'zip\')">下载菜品图包（zip·'+r.imgCount+'张）</button>');
    parts.push('<span style="font-size:11px;opacity:.7;">数据已自动保存，返回网站即可自动导入</span>');
    banner(parts.join(' '),'#F0FDF4');
  }
  
  /* v3.4：返回餐谋长网站函数 */
  window.__cmzReturnToSite = function(){
    /* 跳转到餐谋长网站，并带上参数表示有新抓取的数据 */
    window.open('https://2169913411-afk.github.io/?scraped=1', '_blank');
  };
  (async function(){
    if(isJDMerchantPage()){ await runJDMerchant(); return; }
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
