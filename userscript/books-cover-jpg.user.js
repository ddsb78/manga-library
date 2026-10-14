// ==UserScript==
// @name         书库采集 · 金石堂 / 三民 / 豆瓣 / 出版社资料与封面
// @namespace    local.books.cover
// @version      2.8.0
// @description  勾选或粘贴本站链接，核对作品与卷号，导出作品／单卷原图 ZIP 或 Excel；新版 ZIP 需 PC 1.0.12。支持读册生活、金石堂、三民、豆瓣、出版社及博客来。
// @author       you
// @homepageURL  https://github.com/ddsb78/manga-library/tree/main/userscript
// @updateURL    https://raw.githubusercontent.com/ddsb78/manga-library/main/userscript/books-cover-jpg.user.js
// @downloadURL  https://raw.githubusercontent.com/ddsb78/manga-library/main/userscript/books-cover-jpg.user.js
// @match        https://www.books.com.tw/products/*
// @match        https://activity.books.com.tw/crosscat/show/*
// @match        https://search.books.com.tw/search/query/*
// @match        https://www.kingstone.com.tw/basic/*
// @match        https://www.kingstone.com.tw/search/key/*
// @match        https://www.sanmin.com.tw/product/index/*
// @match        https://www.sanmin.com.tw/search*
// @match        https://www.eslite.com/product/*
// @match        https://www.eslite.com/Search*
// @match        https://www.eslite.com/search*
// @match        https://www.taaze.tw/products/*
// @match        https://www.taaze.tw/rwd_searchResult.html*
// @match        https://book.douban.com/subject/*
// @match        https://book.douban.com/series/*
// @match        https://www.tongli.com.tw/BooksDetail.aspx*
// @match        https://www.egmanga.com.tw/products/*
// @match        https://www.tohan.com.tw/product.php*
// @match        https://www.spp.com.tw/SalePage/Index/*
// @match        https://www.ching-win.com.tw/product-detail/*
// @match        https://www.ching-win.com.tw/products/*
// @match        https://www.kadokawa.com.tw/products/*
// @match        https://jd-intl.com/product/*
// @match        https://jd-intl.com/product-category/*
// @match        https://www.dalapub.com/Publication/*
// @match        https://future-digi.com/index.php*
// @match        https://www.citicomics-shop.com/products*
// @match        https://www.citicomics-shop.com/categories/*
// @match        https://www.s-manga.net/items/contents.html*
// @match        https://www.s-manga.net/search/search.html*
// @match        https://www.shogakukan.co.jp/books/*
// @match        https://www.kodansha.co.jp/comic/products/*
// @match        https://www.kodansha.co.jp/titles/*
// @match        https://www.d-art-shop.tw/products/*
// @match        https://www.d-art-shop.tw/collections/*
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @connect      www.books.com.tw
// @connect      cart.books.com.tw
// @connect      im1.book.com.tw
// @connect      im2.book.com.tw
// @connect      www.kingstone.com.tw
// @connect      cdn.kingstone.com.tw
// @connect      www.sanmin.com.tw
// @connect      cdnec.sanmin.com.tw
// @connect      www.eslite.com
// @connect      s.eslite.com
// @connect      s2.eslite.com
// @connect      www.taaze.tw
// @connect      media.taaze.tw
// @connect      book.douban.com
// @connect      img.doubanio.com
// @connect      img1.doubanio.com
// @connect      img2.doubanio.com
// @connect      img3.doubanio.com
// @connect      img9.doubanio.com
// @connect      www.tongli.com.tw
// @connect      www.egmanga.com.tw
// @connect      cdn-general.cybassets.com
// @connect      www.tohan.com.tw
// @connect      www.spp.com.tw
// @connect      img.91app.com
// @connect      www.ching-win.com.tw
// @connect      www.kadokawa.com.tw
// @connect      img.shoplineapp.com
// @connect      shoplineimg.com
// @connect      jd-intl.com
// @connect      www.dalapub.com
// @connect      future-digi.com
// @connect      www.citicomics-shop.com
// @connect      www.s-manga.net
// @connect      dosbg3xlm0x1t.cloudfront.net
// @connect      www.shogakukan.co.jp
// @connect      www.kodansha.co.jp
// @connect      dvs-cover.kodansha.co.jp
// @connect      www.d-art-shop.tw
// @connect      cdn.store-assets.com
// @run-at       document-idle
// @noframes
// ==/UserScript==

// Operate: extend the existing compact dark toolbox. Read the current page,
// review selected volumes, then choose search XLSX or local-library ZIP.
// Page data is untrusted text, never executable markup in the toolbox.
(function () {
  'use strict';
  // Publishers use isolated adapters; the export schema and review flow remain shared.
  function publisherAdaptersTW() {
    const text = (doc, selector) => clean(doc.querySelector(selector)?.textContent);
    const meta = (doc, name) => doc.querySelector(`meta[property="${name}"]`)?.content || '';
    const fields = (nodes) => {
      const out = {};
      for (const node of nodes) {
        const value = clean(node.textContent), match = value.match(/^([^：:]{1,18})[：:]\s*(.*)$/);
        if (match) out[match[1].replace(/\s/g, '')] = match[2];
      }
      return out;
    };
    const authors = value => clean(value).split(/\s*[/／、;；]\s*/).filter(Boolean);
    const date = value => {
      const m = clean(value).match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
      return m ? `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}` : null;
    };
    const amount = value => { const m = clean(value).match(/\d[\d,]*(?:\.\d+)?/); return m ? Number(m[0].replaceAll(',','')) : null; };
    const paragraphs = node => {
      if (!node) return null;
      const copy = node.cloneNode(true);
      copy.querySelectorAll('script,style').forEach(n=>n.remove());
      copy.querySelectorAll('br').forEach(n=>n.replaceWith('\n'));
      copy.querySelectorAll('p,div,li').forEach(n=>{n.prepend('\n');n.append('\n');});
      return copy.textContent.split('\n').map(clean).filter(Boolean).join('\n') || null;
    };
    const abs = (value, base) => { if(!value) return '';try {return new URL(value,base).href;} catch {return '';} };
    const defaults = (title,publisher,currency='TWD') => ({title,authors:[],publisher,isbn:null,legacy_barcode:null,pub_date:null,description:null,spec:'',rating:'',list_price:null,currency,cover:''});
    const simpleProduct = (host, pattern) => u => u.hostname===host && pattern.test(u.pathname) ? `https://${host}${u.pathname.replace(/\/$/,'')}` : '';
    const links = (doc, selector, base, product) => {
      const found = new Map();
      for (const a of doc.querySelectorAll(selector)) {
        try {
          const url=product(new URL(a.getAttribute('href'),base));
          const title=clean(a.querySelector('img')?.alt || a.textContent);
          if (url && title && !found.has(url)) found.set(url,{url,title});
        } catch { /* Invalid page links are not book records. */ }
      }
      return [...found.values()];
    };
    const detailOnly = {list:()=>false,collect:()=>[],series:()=>''};
    const tongliProduct = u => {const id=u.searchParams.get('BD') || u.searchParams.get('Bd') || '';return u.hostname==='www.tongli.com.tw' && /^\/BooksDetail\.aspx$/i.test(u.pathname) && /^[A-Za-z0-9_-]+$/.test(id) ? `https://www.tongli.com.tw/BooksDetail.aspx?BD=${id}` : '';};
    const chingProduct = u => u.hostname==='www.ching-win.com.tw' && /^\/product-detail\/[A-Za-z0-9_-]+(?:\/[^/]+)?\/?$/.test(u.pathname) ? `https://www.ching-win.com.tw/product-detail/${u.pathname.split('/')[2]}` : '';
    const jdProduct = simpleProduct('jd-intl.com',/^\/product\/[^/]+\/?$/);
    return [
      {key:'tongli',label:'东立',region:'TW',hosts:['www.tongli.com.tw'],coverHosts:['www.tongli.com.tw'],product:tongliProduct,
        list:u=>u.hostname==='www.tongli.com.tw' && /^\/BooksDetail\.aspx$/i.test(u.pathname),
        parse(doc,url) {
          const title=text(doc,'#ContentPlaceHolder1_CBookName') || meta(doc,'og:title').replace(/\s+-\s+.*$/,'');
          const raw=defaults(title,'东立');
          const read=name=>text(doc,'#ContentPlaceHolder1_'+name);
          const volume=read('SEQ').match(/第\s*(\d+)\s*集/);
          Object.assign(raw,{authors:authors(read('AuthorName')),isbn:read('ISBN'),pub_date:date(read('UplineDate')),list_price:amount(read('Price')),
            rating:read('GradeTxt'),cover:abs(doc.querySelector('#ContentPlaceHolder1_Image1')?.getAttribute('src') || meta(doc,'og:image'),url),
            description:paragraphs(doc.querySelector('#ContentPlaceHolder1_Description'))});
          if(volume) raw.volume_number=volume[1];
          return raw;
        },collect:(doc,url)=>links(doc,'#owl-demo2 a[href]',url,tongliProduct),series:()=>''},
      {key:'egmanga',label:'长鸿',region:'TW',hosts:['www.egmanga.com.tw'],coverHosts:['cdn-general.cybassets.com','www.egmanga.com.tw'],product:simpleProduct('www.egmanga.com.tw',/^\/products\/[A-Za-z0-9_-]+\/?$/),...detailOnly,
        parse(doc,url) {
          const f=fields(doc.querySelectorAll('.product_brief p'));
          const raw=defaults(text(doc,'h1') || meta(doc,'og:title'),'长鸿');
          Object.assign(raw,{authors:authors(f['作者']),isbn:f.ISBN,legacy_barcode:f['商品條碼'],pub_date:date(f['出版日期']),spec:f['開本'] || '',rating:f['分級'] || '',
            list_price:amount(doc.querySelector('.product_price del .money_tag')?.getAttribute('data-amount') || text(doc,'.product_price del')),
            cover:meta(doc,'og:image')});
          // CYBERBIZ supplies the larger gallery URL explicitly; never manufacture it.
          for(const script of doc.querySelectorAll('script')) {
            if(!/renderGallery\s*\(/.test(script.textContent)) continue;
            const supplied=script.textContent.match(/\burls\s*:\s*(\[(?:\s*"(?:[^"\\]|\\.)*"\s*,?)*\])/);
            if(supplied) {try {raw.cover=abs(JSON.parse(supplied[1])[0] || raw.cover,url);} catch { /* Ignore malformed optional page metadata. */ }}
          }
          const intro=doc.querySelector('.ckeditor');
          if(intro) {
            const copy=intro.cloneNode(true); let started=false;
            for(const child of [...(copy.querySelector('div') || copy).children]) {
              if(/^故事簡介$/.test(clean(child.textContent))) {started=true;child.remove();}
              else if(!started) child.remove();
              else if(/電子書這邊買/.test(child.textContent)) child.remove();
            }
            raw.description=started ? paragraphs(copy) : null;
          }
          return raw;
        }},
      {key:'tohan',label:'台湾东贩',region:'TW',hosts:['www.tohan.com.tw'],coverHosts:['www.tohan.com.tw'],...detailOnly,
        product:u=>u.hostname==='www.tohan.com.tw' && u.pathname==='/product.php' && u.searchParams.get('act')==='view' && /^\d+$/.test(u.searchParams.get('id') || '') ? `https://www.tohan.com.tw/product.php?act=view&id=${u.searchParams.get('id')}` : '',
        parse(doc,url) {
          const f={};for(const row of doc.querySelectorAll('.col.details .book-row')) f[text(row,'.tit')]=text(row,'.content');
          const raw=defaults(text(doc,'.book-name'),'东贩');
          return Object.assign(raw,{authors:authors(f['作者']).map(name=>name.replace(/^(?:原作|漫畫)\s*/,'')),isbn:f.ISBN,pub_date:date(f['出版日期']),list_price:amount(f['定價']),
            cover:abs(doc.querySelector('.product-slider .slider-for a[href]')?.getAttribute('href') || meta(doc,'og:image'),url),
            description:paragraphs(doc.querySelector('#menu1')),spec:text(doc,'#menu3')});
        }},
      {key:'spp',label:'尖端',region:'TW',hosts:['www.spp.com.tw'],coverHosts:['img.91app.com'],product:simpleProduct('www.spp.com.tw',/^\/SalePage\/Index\/\d+\/?$/i),...detailOnly,
        parse(doc) {
          const source=[...doc.querySelectorAll('script')].map(n=>n.textContent).join('\n');
          const value=key=> {const m=source.match(new RegExp('"'+key+'"\\s*:\\s*("(?:[^"\\\\]|\\\\.)*"|[\\d.]+)'));try{return m ? JSON.parse(m[1]) : null;}catch{return null;}};
          const fragment=new DOMParser().parseFromString(value('ShortDescription') || '', 'text/html');
          const f=fields(fragment.querySelectorAll('li'));
          let name='';for(const script of doc.querySelectorAll('script[type="application/ld+json"]')) {try{const ld=JSON.parse(script.textContent);if(ld['@type']==='Product') name=ld.name || '';}catch{ /* Ignore malformed optional page metadata. */ }}
          const raw=defaults(value('Title') || name,'尖端');
          const suppliedCover=[...doc.querySelectorAll('meta[property="og:image"]')].map(n=>n.content).find(value=>/^https:\/\//.test(value)) || '';
          return Object.assign(raw,{authors:authors(f['作者']),isbn:f['條碼'],pub_date:date(f['上市日']),spec:f['規格'] || '',rating:f['等級'] || '',list_price:value('SuggestPrice'),cover:suppliedCover});
        }},
      {key:'chingwin',label:'青文',region:'TW',hosts:['www.ching-win.com.tw'],coverHosts:['www.ching-win.com.tw'],product:chingProduct,
        list:u=>u.hostname==='www.ching-win.com.tw' && (/^\/products\/?$/.test(u.pathname) && /^[A-Za-z0-9_-]+$/.test(u.searchParams.get('First') || '') || /^\/products\/chingwin\/books(?:\/[A-Za-z0-9_-]+)*\/?$/.test(u.pathname)),
        parse(doc,url) {
          const f=fields(doc.querySelectorAll('.linebox > div,.linebox .line-inbox'));
          const raw=defaults(text(doc,'.productTitle .title') || text(doc,'h1') || meta(doc,'og:title').split(' - ')[0],'青文');
          return Object.assign(raw,{authors:authors(f['作家']),isbn:f.ISBN,pub_date:date(f['出版日期']),spec:f['開本'] || f['尺寸'] || '',rating:f['級別'] || '',
            cover:abs(meta(doc,'og:image'),url),description:paragraphs(doc.querySelector('.tag-content .item-2'))});
        },collect:(doc,url)=>links(doc,'.productBox .box-front .nameBox a.title',url,chingProduct),
        series(doc,url) {const a=[...doc.querySelectorAll('.linebox a[href]')].find(a=>/第\d+集|全系列/.test(a.textContent) && /[?&]First=/.test(a.getAttribute('href')));return a ? abs(a.getAttribute('href'),url) : '';}},
      {key:'kadokawatw',label:'台湾角川',region:'TW',hosts:['www.kadokawa.com.tw'],coverHosts:['img.shoplineapp.com','shoplineimg.com'],product:simpleProduct('www.kadokawa.com.tw',/^\/products\/[^/]+\/?$/),...detailOnly,
        parse(doc) {
          const f=fields((paragraphs(doc.querySelector('.Product-summary')) || '').split('\n').map(value=>({textContent:value})));
          const raw=defaults(text(doc,'.Product-title') || text(doc,'h1'),'台湾角川');
          return Object.assign(raw,{authors:authors(f['作者資訊']),isbn:f.ISBN,pub_date:date(f['上市日期']),list_price:amount(text(doc,'.price-regular.js-price')),cover:meta(doc,'og:image')});
        }},
      {key:'jdintl',label:'玉皇朝国际',region:'HK',hosts:['jd-intl.com'],coverHosts:['jd-intl.com'],product:jdProduct,
        list:u=>u.hostname==='jd-intl.com' && /^\/product-category\/[^/]+\/(?:page\/\d+\/)?$/.test(u.pathname),
        parse(doc,url) {
          const node=doc.querySelector('.woocommerce-product-details__short-description'), content=paragraphs(node) || '';
          const f=fields(content.split('\n').map(value=>({textContent:value})));
          const raw=defaults(text(doc,'h1.product_title') || text(doc,'h1'),'玉皇朝','HKD');
          const volume=raw.title.match(/\s*第\s*(\d+)\s*期(?:\s*完)?\s*$/);
          if(volume) {raw.volume_number=volume[1];raw.work_title=raw.title.slice(0,volume.index).replace(/豪華版|完全版|珍藏版|愛藏版|新裝版/g,'').trim();}
          return Object.assign(raw,{authors:[f['原作'],f['腳本'],f['作畫']].filter(Boolean),isbn:(content.match(/ISBN\s*[:：]?\s*([\d-]+)/i)||[])[1] || null,
            list_price:amount((content.match(/HKD\s*[\d,.]+/i)||[])[0]),cover:abs(doc.querySelector('.woocommerce-product-gallery__image a[href]')?.getAttribute('href') || meta(doc,'og:image'),url)});
        },collect:(doc,url)=>links(doc,'ul.products > li.product a.woocommerce-LoopProduct-link',url,jdProduct),
        series:(doc,url)=>abs(doc.querySelector('.product_meta .posted_in a[href*="/product-category/"]')?.getAttribute('href') || '',url)},
    ];
  }
  function publisherAdaptersJP() {
    const text=(doc,selector)=>clean(doc.querySelector(selector)?.textContent);
    const meta=(doc,name)=>doc.querySelector(`meta[property="${name}"]`)?.content || '';
    const abs=(value,base)=>{try{return value ? new URL(value,base).href : '';}catch{return '';}};
    const date=value=>{const m=String(value || '').match(/(\d{4})[年./-](\d{1,2})[月./-](\d{1,2})/);return m ? `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}` : '';};
    const amount=value=>{const m=String(value || '').match(/\d[\d,]*(?:\.\d+)?/);return m ? Number(m[0].replaceAll(',','')) : null;};
    const paragraphs=node=>{if(!node)return '';const copy=node.cloneNode(true);copy.querySelectorAll('script,style').forEach(n=>n.remove());copy.querySelectorAll('br').forEach(n=>n.replaceWith('\n'));copy.querySelectorAll('p,div').forEach(n=>{n.prepend('\n');n.append('\n');});return copy.textContent.split('\n').map(clean).filter(Boolean).join('\n');};
    const htmlText=(doc,value)=>{const node=doc.createElement('div');node.innerHTML=String(value || '');return paragraphs(node);};
    const defaults=(title,publisher,currency='JPY')=>({title,authors:[],publisher,isbn:'',legacy_barcode:'',pub_date:'',description:'',spec:'',rating:'',list_price:null,currency,cover:''});
    const product=(host,pattern)=>u=>u.hostname===host && pattern.test(u.pathname) ? `https://${host}${u.pathname.replace(/\/$/,'')}` : '';
    const smProduct=u=>{
      if(u.hostname!=='www.s-manga.net' || u.pathname!=='/items/contents.html')return '';
      const isbn=u.searchParams.get('isbn');
      return isbn && /^\d[\d-]{9,20}$/.test(isbn) ? `https://www.s-manga.net/items/contents.html?isbn=${isbn}` : '';
    };
    const shoProduct=product('www.shogakukan.co.jp',/^\/books\/\d{8}\/?$/);
    const kodProduct=product('www.kodansha.co.jp',/^\/comic\/products\/\d+\/?$/);
    const dartProduct=u=>u.hostname==='www.d-art-shop.tw' && /^\/(?:collections\/[^/]+\/)?products\/[^/]+\/?$/.test(u.pathname) ? `https://www.d-art-shop.tw/products/${u.pathname.split('/').filter(Boolean).at(-1)}` : '';
    const links=(doc,selector,url,canonical)=>{
      const found=new Map();for(const a of doc.querySelectorAll(selector)){let value='';try{value=canonical(new URL(a.getAttribute('href'),url));}catch{ /* Ignore malformed optional page metadata. */ }if(value && !found.has(value))found.set(value,{url:value,title:clean(a.textContent) || clean(a.querySelector('img')?.alt)});}return [...found.values()];
    };
    const inertia=doc=>{try{return JSON.parse(doc.querySelector('[data-page]')?.getAttribute('data-page') || '{}').props || {};}catch{return {};}};
    return [
      {key:'smanga',label:'集英社 S-MANGA',region:'JP',hosts:['www.s-manga.net'],coverHosts:['dosbg3xlm0x1t.cloudfront.net'],product:smProduct,
        list:u=>u.hostname==='www.s-manga.net' && u.pathname==='/search/search.html' && /^\d+$/.test(u.searchParams.get('seriesid') || '') && !u.searchParams.get('digital'),
        parse(doc){
          const raw=defaults(text(doc,'h1.bktitle') || text(doc,'.item-info h1'),'集英社');
          const paper=[...doc.querySelectorAll('.item-info-kamidigi li section')].find(n=>text(n,'h4')==='紙版');
          const info=clean(paper?.textContent), isbn=text(paper || doc,'#item-ISBN').replace(/^ISBN\s*[:：]\s*/i,'');
          raw.authors=[...doc.querySelectorAll('.item-info-author a')].map(n=>clean(n.textContent));
          return Object.assign(raw,{isbn,pub_date:date(info),list_price:amount((info.match(/[\d,]+円/) || [])[0]),spec:(info.match(/([^\s]+判)／/) || [])[1] || '',description:paragraphs(doc.querySelector('.item-introduction p')),cover:meta(doc,'og:image'),blocked:!paper || !isbn ? '未找到纸版书目资料，请核对实体版本' : ''});
        },collect:(doc,url)=>links(doc,'.single-item a.searchresult-a[href*="isbn="],.single-item .bktitle a[href*="isbn="]',url,smProduct),
        series:(doc,url)=>abs(doc.querySelector('.item-btn-zenkan a[href*="seriesid="]')?.getAttribute('href'),url)},
      {key:'shogakukan',label:'小学馆',region:'JP',hosts:['www.shogakukan.co.jp'],coverHosts:['www.shogakukan.co.jp'],product:shoProduct,
        list:u=>u.hostname==='www.shogakukan.co.jp' && /^\/books\/volume\/\d+\/?$/.test(u.pathname),
        parse(doc,url){
          const book=inertia(doc).book || {}, raw=defaults(clean(book.title),'小学馆');
          return Object.assign(raw,{work_title:clean(book.daihyo_shohin?.daihyo_name),volume_number:book.kansu ? String(book.kansu) : '',authors:(book.auth_notations || []).map(a=>clean(a.auth_name)),isbn:book.isbn13_cd || book.isbn_cd || '',pub_date:date(book.publish_date),list_price:book.teika_now ?? null,spec:book.hankei || '',description:htmlText(doc,book.promo_contents?.[3]),cover:abs(book.image_url,url),blocked:!book.isbn_cd || book.product_type==='digital' ? '未找到实体书目资料，请核对版本' : ''});
        },collect(doc,url){const rendered=links(doc,'#main .bookBox01 p.ttl a[href]',url,shoProduct);if(rendered.length)return rendered;const books=inertia(doc).books, rows=Array.isArray(books)?books:books?.data || [];return rows.flatMap(row=>{const b=row.document?.content || row;let canonical='';try{canonical=shoProduct(new URL(b.url,url));}catch{ /* Ignore malformed optional page metadata. */ }return canonical?[{url:canonical,title:clean(b.title)}]:[];});},
        series(doc){const book=inertia(doc).book;return book?.has_zenkan && book?.daihyo_cd ? `https://www.shogakukan.co.jp/books/volume/${book.daihyo_cd}` : ''; }},
      {key:'kodansha',label:'讲谈社',region:'JP',hosts:['www.kodansha.co.jp'],coverHosts:['dvs-cover.kodansha.co.jp'],product:kodProduct,
        list:u=>u.hostname==='www.kodansha.co.jp' && /^\/titles\/\d+\/?$/.test(u.pathname),
        parse(doc){
          const raw=defaults(text(doc,'h1.font-bold.text-kds-20'),'讲谈社');
          const heading=[...doc.querySelectorAll('h4')].find(n=>clean(n.textContent)==='紙版'), fields={};
          for(const row of heading?.parentElement?.querySelectorAll('section') || []){const label=text(row,'h5');if(label)fields[label]=text(row,'p');}
          const title=doc.querySelector('h1.font-bold.text-kds-20');
          raw.authors=[...title?.parentElement?.querySelectorAll('li a[href*="/products/search?"]') || []].map(n=>clean(n.textContent));
          const intro=[...doc.querySelectorAll('p')].find(n=>n.querySelector('br') && clean(n.textContent).length>100 && !n.closest('footer'));
          return Object.assign(raw,{isbn:fields.ISBN || '',pub_date:date(fields['発売日']),spec:fields['判型'] || '',list_price:amount(fields['価格']),description:paragraphs(intro),cover:meta(doc,'og:image'),blocked:!heading || !fields.ISBN ? '电子书或未找到纸版资料，请核对实体版本' : ''});
        },collect:(doc,url)=>links(doc,'[data-scroll-anchor="publications"] a[href*="/comic/products/"]',url,kodProduct),
        series(doc,url){const a=[...doc.querySelectorAll('a[href*="/titles/"]')].find(n=>/既刊一覧/.test(n.textContent));return abs(a?.getAttribute('href'),url);}},
      {key:'dart',label:'d/art 原动力视觉',region:'TW',hosts:['www.d-art-shop.tw'],coverHosts:['cdn.store-assets.com'],product:dartProduct,
        list:u=>u.hostname==='www.d-art-shop.tw' && /^\/collections\/[^/]+\/?$/.test(u.pathname),
        parse(doc,url){
          const raw=defaults(text(doc,'.product-single h1'),'原动力视觉','TWD'), node=doc.querySelector('.product-description.rte'), content=paragraphs(node);
          const field=label=>clean((content.match(new RegExp(`${label}\\s*[｜:：]\\s*([^\\n]+)`)) || [])[1]);
          const cover=[...node?.querySelectorAll('img') || []].find(n=>/(?:^|[_\s])封面$/.test(clean(n.alt)));
          const category=field('商品類別'), company=field('發行公司'), title=raw.title;
          const sku=field('商品編號');
          return Object.assign(raw,{authors:field('作\\s*者').split(/[、／]/).map(clean).filter(Boolean),isbn:field('ISBN'),legacy_barcode:field('條\\s*碼'),pub_date:date(field('出版日期') || field('發行日期')),spec:(content.match(/[AB]\d開本[^\n]*/) || [])[0] || '',rating:/18禁|限制級/.test(content)?'限制級':'',list_price:amount(field('售\\s*價')),cover:abs(cover?.getAttribute('src'),url),publisher:/原動力視覺|d\/art/i.test(company)?'原动力视觉':company || null,
            blocked:category && category!=='漫畫' && category!=='畫集' ? '此商品不是可导出的书籍，请核对商品类别' : /套組|合售|套書/.test(title) ? '特典套组或合售不能当作一卷导出，请核对单本' : !/^DAV-/i.test(sku) || !/繁中|繁體中文/.test(title+' '+field('商品版本')) ? '仅支持已确认的 d/art 繁中出版单本，请核对发行公司与版本' : ''});
        },collect:(doc,url)=>[...doc.querySelectorAll('.grid-uniform.grid-link__container .product_grid-item')].map(row=>{const a=row.querySelector('a[href*="/products/"]');return {url:a ? dartProduct(new URL(a.getAttribute('href'),url)) : '',title:text(row,'.grid-link__title')};}).filter(row=>row.url && row.title),series:()=>''},
    ];
  }
  function publisherAdaptersExtra() {
    const txt = (doc, selector) => clean(doc.querySelector(selector)?.textContent);
    const meta = (doc, name) => doc.querySelector(`meta[property="${name}"]`)?.content || '';
    const lines = node => node ? String(node.innerText || node.textContent || '').trim() : '';
    const raw = publisher => ({title:'', authors:[], publisher, isbn:'', legacy_barcode:'',
      pub_date:'', description:'', spec:'', rating:'', list_price:'', currency:'', cover:''});
    const field = (text, label) => clean((text.match(new RegExp(`(?:^|[\\n\\r])[ \\t\u3000]*${label}[ \\t\u3000]*[：:][ \\t\u3000]*([^\\n\\r]+)`, 'i')) || [])[1]);
    const section = (doc, heading) => {
      const node = [...doc.querySelectorAll('h2')].find(n => clean(n.textContent) === heading);
      return node?.parentElement?.nextElementSibling || null;
    };
    const jsonProducts = doc => {
      const nodes = [];
      const visit = value => {
        if (Array.isArray(value)) return value.forEach(visit);
        if (!value || typeof value !== 'object') return;
        const types = [].concat(value['@type'] || []);
        if (types.includes('Product') || types.includes('Book')) nodes.push(value);
        if (value['@graph']) visit(value['@graph']);
      };
      for (const script of doc.querySelectorAll('script[type="application/ld+json"]')) {
        try { visit(JSON.parse(script.textContent)); } catch { /* malformed page metadata */ }
      }
      return nodes;
    };
    const collect = (doc, url, product, selector) => {
      const result = new Map();
      for (const a of doc.querySelectorAll(selector)) {
        try {
          const target = product(new URL(a.getAttribute('href') || '', url));
          const title = clean(a.textContent || a.querySelector('img')?.alt);
          if (target && title && !result.has(target)) result.set(target, {url:target, title});
        } catch { /* invalid catalog link */ }
      }
      return [...result.values()];
    };
    const dalaProduct = u => u.hostname === 'www.dalapub.com' && u.pathname === '/Publication/Detail' &&
      /^[A-Za-z0-9 _-]+$/.test(u.searchParams.get('code') || '')
      ? 'https://www.dalapub.com/Publication/Detail?code=' + encodeURIComponent(u.searchParams.get('code')) : '';
    const citiProduct = u => {
      if(u.hostname!=='www.citicomics-shop.com' || !/^\/products\/[^/]+\/?$/.test(u.pathname)) return '';
      const variant=u.searchParams.get('bcd_variant');
      if(variant!==null && !/^[a-f0-9]{24}$/.test(variant)) return '';
      return 'https://www.citicomics-shop.com'+u.pathname.replace(/\/$/,'')+(variant?'?bcd_variant='+variant:'');
    };
    const futureProduct = u => u.hostname === 'future-digi.com' && u.pathname === '/index.php' &&
      u.searchParams.get('route') === 'product/product' && /^\d+$/.test(u.searchParams.get('product_id') || '')
      ? 'https://future-digi.com/index.php?route=product/product&product_id=' + u.searchParams.get('product_id') : '';
    return [{
      key:'dala', label:'大辣', region:'TW', hosts:['www.dalapub.com'], coverHosts:['www.dalapub.com'],
      product:dalaProduct,
      list:u => u.hostname === 'www.dalapub.com' && u.pathname === '/Publication/List',
      parse(doc) {
        const out = raw('大辣');
        out.title = txt(doc, '.member-container .info-group .title');
        const info = section(doc, '書籍資訊');
        // BRs are field boundaries even when DOMParser has no rendered innerText.
        const copy = info?.cloneNode(true);
        copy?.querySelectorAll('br').forEach(n => n.replaceWith('\n'));
        const text = lines(copy);
        out.authors = [field(text, '作者') || txt(doc, '.member-group a.name').replace(/^作者[：:]\s*/, '')].filter(Boolean);
        out.isbn = field(text, 'ISBN');
        out.pub_date = field(text, '出版日期');
        out.spec = [field(text, '規格'), field(text, '頁數'), field(text, '裝訂'), field(text, '顏色')].filter(Boolean).join(' / ');
        out.description = lines(section(doc, '書籍介紹'));
        out.cover = doc.querySelector('.member-container > .container > .img-container > img')?.getAttribute('src') || '';
        if (!out.title || !out.isbn) out.blocked = '未取得单本书籍资料';
        return out;
      },
      collect:(doc,url) => [...doc.querySelectorAll('main .detail-group a[href*="/Publication/Detail"]')].map(a=>({url:dalaProduct(new URL(a.getAttribute('href'),url)),title:txt(a.closest('.detail-group'),'.book').replace(/^書名\s*\|\s*/,'')})).filter(row=>row.url && row.title),
      series:() => ''
    }, {
      key:'citicomics', label:'文化傳信', region:'HK', hosts:['www.citicomics-shop.com'],
      coverHosts:['img.shoplineapp.com','shoplineimg.com'], product:citiProduct,
      list:u => u.hostname === 'www.citicomics-shop.com' && (u.pathname === '/products' || /^\/categories\//.test(u.pathname)),
      parse(doc,url) {
        const out = raw('文化傳信');
        if(new URL(url).searchParams.has('bcd_variant')) return citiVariantRaw(doc,url);
        const data = jsonProducts(doc).find(v => clean(v.name) === txt(doc,'h1.Product-title'));
        out.title = txt(doc,'h1.Product-title');
        const summary = doc.querySelector('.Product-summary, .Product-shortDescription, .product-summary');
        const copy = summary?.cloneNode(true);
        copy?.querySelectorAll('br').forEach(n => n.replaceWith('\n'));
        const text = lines(copy);
        out.authors = [field(text,'原作') || field(text,'作者')].filter(Boolean);
        out.isbn = field(text,'ISBN');
        out.spec = field(text,'尺寸');
        out.rating = field(text,'級別');
        out.description = clean(data?.description);
        const priceText = field(text,'售價') || clean((text.match(/(?:^|\n)售價\s+([^\n]+)/) || [])[1]);
        const price = priceText.match(/\d[\d,]*(?:\.\d+)?/);
        out.list_price = price ? Number(price[0].replaceAll(',','')) : null;
        out.currency = 'HKD';
        const image = data?.image;
        out.cover = typeof image === 'string' ? image : Array.isArray(image) ? image[0] : image?.url || meta(doc,'og:image');
        const variants = doc.querySelector('select[name="variation"], select[name="product_variation"], .product-variation select');
        if (!out.title || !field(text,'書名') || !field(text,'期數')) out.blocked = '未取得单本书籍资料';
        if (variants?.options?.length > 1 || /(?:\d+\s*[-~～]\s*\d+|任選|任选|選擇期數|选择期数)/.test(field(text,'期數')))
          out.blocked = '同页包含多卷，请先确认具体期数与封面';
        if (/^全[一1](?:冊|册|卷)$/.test(field(text,'期數')))
          Object.assign(out,{volume_number:'全',volume_sequence_format:'single',published_volume_count:1,serial_status:'completed'});
        return out;
      },
      collect:(doc,url) => [...doc.querySelectorAll('.ProductList-list .product-item a.Product-item[href]')].map(a=>({url:citiProduct(new URL(a.getAttribute('href'),url)),title:txt(a,'.title')})).filter(row=>row.url && row.title),
      series:() => ''
    }, {
      key:'future', label:'未来数位', region:'TW', hosts:['future-digi.com'], coverHosts:['future-digi.com'],
      product:futureProduct,
      list:u => u.hostname === 'future-digi.com' && u.pathname === '/index.php' &&
        ['product/category','product/search','product/manufacturer/info'].includes(u.searchParams.get('route')),
      parse(doc) {
        const gate=doc.querySelector('img[src$="/images/ticrf.jpg"]');
        if(gate && !gate.closest('[hidden],[style*="display: none"],[style*="display:none"]'))
          throw failure('请先在未来数位完成已满十八岁确认，再刷新商品页读取。','login');
        const out=raw('未來數位'), heading=txt(doc.querySelector('#product')?.parentElement || doc,'h1');
        const sections=[...doc.querySelectorAll('#tab-description h3')];
        const text=sections.filter(n=>/編[\s\u3000]*號[：:]|產品規格/.test(n.textContent)).map(n=>{
          const copy=n.cloneNode(true);copy.querySelectorAll('br').forEach(b=>b.replaceWith('\n'));return lines(copy);
        }).join('\n'), read=label=>field(text,label), category=read('類[\\s　]*別');
        out.title=heading.replace(/^\[\d{1,2}\/\d{1,2}\]\s*/,'').replace(/※\d{1,2}\/\d{1,2}出貨$/,'');
        const sourceAuthors=read('作[\\s　]*者');
        out.authors=sourceAuthors.replace(/(?:^|\s)(?:原作|漫畫|作畫|角色原案|腳本原案)\s*[：:／]\s*/g,'、').split(/[、／]/).map(clean).filter(Boolean);
        out.publisher=txt(doc,'#content a[href*="route=product/manufacturer/info"]') || out.publisher;
        out.isbn=read('ISBN'); out.legacy_barcode=read('條[\\s　]*碼') || read('EAN');
        const date=read('發售日').match(/(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?/);
        out.pub_date=date ? `${date[1]}-${date[2].padStart(2,'0')}-${date[3].padStart(2,'0')}` : '';
        out.spec=[read('大小'),read('頁數') && read('頁數')+'頁'].filter(Boolean).join(' / ');
        const price=read('售[\\s　]*價').match(/\d[\d,]*(?:\.\d+)?/);
        out.list_price=price ? Number(price[0].replaceAll(',','')) : null; out.currency='TWD';
        const ratingText=sections.map(n=>n.textContent).join(' ');
        out.rating=/成人漫畫|18禁|限制級/.test(category+' '+ratingText)?'限制級':/輔導級/.test(ratingText)?'輔導級':/普遍級/.test(ratingText)?'普遍級':'';
        out.cover=doc.querySelector('#content .thumbnails > li:first-child a.thumbnail')?.getAttribute('href') || '';
        const description=doc.querySelector('#tab-description')?.cloneNode(true);
        description?.querySelectorAll('img,script,style').forEach(n=>n.remove());
        for(const node of description?.querySelectorAll('h3') || []) {
          if(/編[\s\u3000]*號[：:]|產品規格|^\s*目錄[：:]/.test(node.textContent)) node.remove();
        }
        description?.querySelectorAll('br').forEach(n=>n.replaceWith('\n'));
        out.description=lines(description).split('\n').map(clean).filter(Boolean).join('\n');
        out.metadata_note='官网商品编号：'+read('編[\\s　]*號')+'；原站作者：'+sourceAuthors+(out.title!==heading?'；原站标题：'+heading:'');
        if(!out.title || !category || !/漫畫|小說|畫[冊集]|書籍|繪本/.test(category))
          out.blocked='未取得单本书籍类别；周边、游戏不能作为卷册导出';
        return out;
      },
      collect:(doc,url) => collect(doc,url,futureProduct,'#content .product-thumb h4 a[href]'),
      series:() => ''
    }];
  }
  function sanminCover(code) {
    const value=clean(code);
    return /^[0-9]{13}$/.test(value) ? `https://cdnec.sanmin.com.tw/product_images/${value.slice(3,6)}/${value.slice(3,12)}.jpg` : '';
  }
  function retailerAdapters() {
    const text=(doc,selector)=>clean(doc.querySelector(selector)?.textContent);
    const meta=(doc,name)=>doc.querySelector(`meta[property="${name}"]`)?.content || '';
    const valueCode=raw=>{const code=clean(raw).replace(/[\s-]/g,'');return isbn13(code) ? {isbn:code} : {legacy_barcode:/^[0-9]{13}$/.test(code)?code:null};};
    const date=raw=>{const m=clean(raw).match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:null;};
    const amount=raw=>{const m=clean(raw).match(/[0-9][0-9,]*(?:\.[0-9]+)?/);return m?Number(m[0].replaceAll(',','')):null;};
    const paragraphs=node=>{if(!node)return null;const copy=node.cloneNode(true);copy.querySelectorAll('br').forEach(n=>n.replaceWith('\n'));copy.querySelectorAll('p').forEach(n=>n.append('\n'));return copy.textContent.trim();};
    const kingstoneFields=doc=>{
      const result={};
      for(const row of doc.querySelectorAll('.table_2col_deda')) {
        const nodes=[...row.children];for(let i=0;i+1<nodes.length;i+=2)result[clean(nodes[i].textContent)]=clean(nodes[i+1].textContent);
      }
      for(const label of doc.querySelectorAll('.basicunit .title_basic')) {
        const row=label.parentElement.cloneNode(true);row.querySelectorAll('.title_basic,.followfield,.qmkfield').forEach(n=>n.remove());
        result[clean(label.textContent).replace(/[：:]$/,'')]=clean(row.textContent);
      }
      if(result['出版日'])result['出版日期']=result['出版日'];
      return result;
    };
    return [{key:'kingstone',label:'金石堂',region:'TW',hosts:['www.kingstone.com.tw'],coverHosts:['cdn.kingstone.com.tw'],
      product:u=>/^\/basic\/[0-9]{13}\/?$/.test(u.pathname)?`${u.origin}${u.pathname.replace(/\/$/,'')}/`:null,
      list:u=>/^\/search\/key\//.test(u.pathname),
      collect:(doc)=>[...doc.querySelectorAll('.modFilterProd .mod_prod_card .pdnamebox a[href],li.displayunit h3 a[href]')].map(a=>({url:a.getAttribute('href'),title:clean(a.textContent)})),
      parse:doc=>{const f=kingstoneFields(doc),code=f.ISBN || f.EAN || '';const supplied=meta(doc,'og:image');
        return {title:text(doc,'h1'),authors:clean(f['作者']).split(/\s*[/／、;；]\s*/).filter(Boolean),publisher:f['出版社'],...valueCode(code),
          pub_date:date(f['出版日期']),description:paragraphs(doc.querySelector('.pdintro_txt1field.panelCon > span')),
          spec:[f['商品規格'],f['裝訂'],f['頁數']?`${f['頁數']}頁`:''].filter(Boolean).join('；'),rating:f['分級'],
          list_price:amount(text(doc,'.basicunit.price1 s b.sty00')),currency:'TWD',
          cover:/\/(?:restricted|noimage|no_image|noimg_load|nopic|no_pic|default)\.[a-z]+$/i.test(supplied)?sanminCover(code):supplied || sanminCover(code),cover_alternative:sanminCover(code)};},
      series:()=>''},
    {key:'sanmin',label:'三民',region:'TW',hosts:['www.sanmin.com.tw'],coverHosts:['cdnec.sanmin.com.tw'],
      product:u=>/^\/product\/index\/[0-9]+\/?$/.test(u.pathname)?`${u.origin}${u.pathname.replace(/\/$/,'')}`:null,
      list:u=>/^\/search\/?$/.test(u.pathname),
      collect:(doc,url)=>{const cards=[...doc.querySelectorAll('.sProduct[id^="Prod"]')];return cards.map(card=>{const a=card.querySelector('.Info .Title a');return {url:a?.getAttribute('href'),title:clean(a?.querySelector('h3')?.textContent || a?.textContent),html:card.outerHTML,queryCode:cards.length===1?new URL(url).searchParams.get('k'):null};});},
      parse:(doc,url,publicSearchCode)=>{const f={};for(const label of doc.querySelectorAll('li.ga .Bold')) {f[clean(label.textContent)]=clean(label.parentElement.textContent).replace(/^[^：:]*[：:]\s*/,'');}
        const list=doc.querySelector('.sProduct'),card=list || doc;
        const attr=selector=>text(card,selector).replace(/^[^：:]*[：:]\s*/,'');
        const code=f.ISBN13 || f.ISBN || (list?publicSearchCode:'') || '';
        const price=[...doc.querySelectorAll('.col-price .Price > div')].map(n=>clean(n.textContent)).find(v=>/^定價\s*[：:]/.test(v));
        return {title:text(doc,'h1') || text(card,'.Info .Title h3'),authors:clean(f['作者'] || attr('.Attribute .Author')).split(/\s*[/／、;；]\s*/).filter(Boolean),publisher:f['出版社'] || attr('.Attribute .Publisher'),...valueCode(code),
          pub_date:date(f['出版日'] || attr('.Attribute .PublishDate')),description:paragraphs(doc.querySelector('#Intro1')?.parentElement.querySelector('.SectionBody')) || text(card,'.Description'),
          spec:[f['規格'],f['裝訂'],f['版次']?`版次：${f['版次']}`:''].filter(Boolean).join('；'),rating:f['分級'] || text(card,'.Labels'),
          list_price:price?amount(price):null,currency:'TWD',cover:meta(doc,'og:image') || card.querySelector('.Cover img')?.getAttribute('data-src') || card.querySelector('.Cover img')?.getAttribute('src') || sanminCover(code),
          metadata_note:list?'三民公开搜索摘要；详情字段未读取':''};},series:()=>''},
    {key:'eslite',label:'诚品书店',region:'TW',hosts:['www.eslite.com'],coverHosts:['s.eslite.com','s2.eslite.com'],
      product:u=>/^\/product\/[0-9]+\/?$/.test(u.pathname)?`${u.origin}${u.pathname.replace(/\/$/,'')}`:null,
      list:u=>/^\/search\/?$/i.test(u.pathname),
      collect:doc=>[...doc.querySelectorAll('.search-product-block .product-item')].map(row=>{const a=row.matches('a[href]')?row:row.querySelector('a[href*="/product/"]');return {url:a?.getAttribute('href'),title:clean(row.querySelector('.product-name')?.textContent)||a?.getAttribute('title')||a?.querySelector('img')?.alt||clean(a?.textContent)};}),
      parse:doc=>{
        const f={};for(const span of doc.querySelectorAll('.ec-col-12 > span')){const parts=clean(span.textContent).split(' / ');if(parts.length===2)f[parts[0]]=parts[1];}
        const row=selector=>clean(doc.querySelector(selector)?.lastElementChild?.textContent);
        let data={};for(const node of doc.querySelectorAll('script[type="application/ld+json"]'))try{const value=JSON.parse(node.textContent);const values=Array.isArray(value)?value:value['@graph']||[value];const found=values.find(v=>v['@type']==='Product'||v['@type']==='Book');if(found){data=found;break;}}catch{/* Malformed page JSON cannot supply book fields. */}
        const code=f.ISBN13 || f['EAN貨碼'] || data.gtin13 || '';const cover=typeof data.image==='string'?data.image:'';
        return {title:text(doc,'h1') || data.name,authors:row('.books-publication-link-row.author').split(/\s*[/／、;；]\s*/).filter(Boolean),
          publisher:row('.books-publication-link-row.supplier') || data.brand?.name,pub_date:date(row('.books-publication-row.publicDate')),...valueCode(code),
          description:data.description || null,spec:[f['尺寸'],f['裝訂'],f['頁數']?`${f['頁數']}頁`:''].filter(Boolean).join('；'),
          rating:/^R:/.test(f['級別']||'')?'限制級':/^N:/.test(f['級別']||'')?'普通級':f['級別'],
          list_price:amount(text(doc,'.price-group .pre-product-price')),currency:'TWD',
          cover:/product-restricted|noimage/i.test(cover)?sanminCover(code):cover,cover_alternative:sanminCover(code),
          metadata_note:'诚品公开书目；动态页面须等待标题、书目与规格加载完成'};
      },series:()=>''},
    {key:'taaze',label:'读册生活 TAAZE',region:'TW',hosts:['www.taaze.tw'],coverHosts:['media.taaze.tw'],
      product:u=>{
        const id=(u.pathname.match(/^\/products\/(111[0-9]{8})\.html$/) || [])[1] || (u.pathname==='/usedList.html' && /^111[0-9]{8}$/.test(u.searchParams.get('oid') || '') ? u.searchParams.get('oid') : '');
        return id ? `${u.origin}/products/${id}.html` : null;
      },
      list:u=>u.pathname==='/rwd_searchResult.html',
      collect:doc=>[...doc.querySelectorAll('#listView .titleMain a[href],#listView a.titleMain[href]')].map(a=>({url:a.getAttribute('href'),title:clean(a.textContent).replace(/\s*\(二手書\)\s*$/,'')})),
      parse:doc=>{
        const f={};for(const node of doc.querySelectorAll('#bottomArea span.prodInfo_boldSpan')) {
          const m=clean(node.textContent).match(/^([^：:]{1,18})[：:]\s*(.*)$/);if(m && !f[m[1]])f[m[1]]=m[2];
        }
        const price=[...doc.querySelectorAll('.price > p')].find(n=>/^定價\s*[：:]/.test(clean(n.textContent)));
        const size=clean(f['開數']);
        return {title:text(doc,'#ga4ProdTitle'),authors:clean(f['作者']).split(/\s*[/／、;；]\s*/).filter(Boolean),publisher:f['出版社'],
          ...valueCode(f['ISBN/ISSN']),pub_date:date(f['出版日期']),description:paragraphs(doc.querySelector('#prodPfDiv')),
          spec:[f['裝訂方式'],f['頁數'],size ? /^\d+$/.test(size) ? `${size}開` : size : ''].filter(Boolean).join('；'),rating:f['分級'],
          list_price:price ? amount(price.textContent) : null,currency:'TWD',
          cover:doc.querySelector('img[src*="media.taaze.tw/showThumbnail.html"]')?.getAttribute('src') || meta(doc,'og:image'),
          metadata_note:f['譯者'] ? '译者：'+f['譯者'] : ''};
      },series:()=>''}];
  }
  const SITE_ADAPTERS = [...publisherAdaptersTW(), ...publisherAdaptersJP(), ...publisherAdaptersExtra(), ...retailerAdapters()];
  const DOUBAN = typeof location !== 'undefined' && location.hostname === 'book.douban.com';
  const CURRENT_SITE = typeof location !== 'undefined' ? SITE_ADAPTERS.find(s=>s.hosts.includes(location.hostname)) : null;
  const SITE_NAME = CURRENT_SITE?.label || (DOUBAN ? '豆瓣' : '博客来');
  const CACHE = CURRENT_SITE ? 'bcd-publisher-v1:'+CURRENT_SITE.key : DOUBAN ? 'bcd-douban-v2' : 'bcd-library-v3';
  const COOLDOWN = CURRENT_SITE ? 'bcd-publisher-cooldown-v1:'+CURRENT_SITE.key : DOUBAN ? 'bcd-douban-cooldown-v1' : 'bcd-cooldown-v2';
  const MAX_ITEMS = 300, MAX_ZIP = 128 * 1024 * 1024, GAP = 2000;
  const SEARCH_CHOICES = {
    kingstone:{rows:'.modFilterProd .mod_prod_card,li.displayunit',link:'.pdnamebox a[href],h3 a[href]'},
    sanmin:{rows:'.sProduct[id^="Prod"]',link:'.Info .Title a[href]',container:'.Info'},
    eslite:{rows:'.search-product-block .product-item',link:'a[href*="/product/"]',container:'.item-wording-wrap'},
    taaze:{rows:'#listView > .media,#listView .bookGridByListView',link:'.titleMain a[href],a.titleMain[href]',container:'.info_frame,.bookCaption'},
  };
  const TTL = 7 * 86400000;
  const EDITIONS = [['tankobon', '单行本'], ['aizouban', '爱藏版'], ['collector', '珍藏版'],
    ['kanzenban', '完全版'], ['full_color', '全彩版'], ['deluxe_collection', '典藏版'],
    ['new_edition', '新装版'], ['bunko', '文库版'], ['luxury', '豪华版']];
  const KINDS = [['main', '正篇'], ['extra', '外传 / 番外'], ['guidebook', '公式书'],
    ['artbook', '画集'], ['novel', '小说'], ['other', '其他附属册']];
  const FORMATS = [['numeric', '数字卷'], ['single', '全一卷'], ['upper_lower', '上下'], ['upper_middle_lower', '上中下']];
  const SERIAL = [['unknown', '待确认'], ['ongoing', '连载中'], ['completed', '已完结']];
  const CONTENTS = [['manga','漫画'],['novel','小说'],['guidebook','公式书'],['artbook','画集']];
  const REGIONS = [['UNKNOWN','待确认'],['HK','港版'],['TW','台版'],['CN','简中'],['JP','日版'],['OTHER','其他']];
  const clean = value => String(value || '').replace(/\s+/g, ' ').trim();
  function publisherRef(id) {
    const match=String(id).match(/^pub-([a-z-]+):(.+)$/);
    if(!match) return null;
    const site=SITE_ADAPTERS.find(s=>s.key===match[1]);
    try {
      const url=new URL(decodeURIComponent(match[2]));
      if(!site || url.protocol!=='https:' || url.username || url.password || url.port || !site.hosts.includes(url.hostname)) return null;
      const canonical=site.product(url);
      return canonical===url.href ? {site,url:canonical} : null;
    } catch {return null;}
  }
  const productUrl = id => id.startsWith('pub-') ? publisherRef(id)?.url || '' : id.startsWith('douban-') ? 'https://book.douban.com/subject/' + id.slice(7) + '/' : 'https://www.books.com.tw/products/' + id;
  const sameSiteId = id => CURRENT_SITE ? publisherRef(id)?.site.key===CURRENT_SITE.key : !id.startsWith('pub-') && id.startsWith('douban-')===DOUBAN;
  function isBooksSearch(value) {
    try {const u=new URL(value);return u.protocol==='https:' && u.hostname==='search.books.com.tw' && u.pathname.startsWith('/search/query/');} catch {return false;}
  }
  function isSearchResults(value) {
    if(isBooksSearch(value)) return true;
    try {
      const u=new URL(value), site=SITE_ADAPTERS.find(s=>s.hosts.includes(u.hostname));
      return u.protocol==='https:' && Boolean(SEARCH_CHOICES[site?.key]) && (site.key==='kingstone' ? u.pathname.startsWith('/search/key/') : site.key==='taaze' ? u.pathname==='/rwd_searchResult.html' : /^\/search\/?$/i.test(u.pathname));
    } catch {return false;}
  }
  function productId(value) {
    try {
      const u = new URL(value, typeof location!=='undefined' && location.protocol==='https:' ? location.href : /^\/subject\//.test(String(value)) ? 'https://book.douban.com' : 'https://www.books.com.tw');
      if (u.protocol !== 'https:' || u.username || u.password || u.port) return '';
      const site=SITE_ADAPTERS.find(s=>s.hosts.includes(u.hostname));
      if(site) {
        const canonical=site.product(u);
        if(!canonical) return '';
        const checked=new URL(canonical);
        if(checked.protocol!=='https:' || checked.username || checked.password || checked.port || !site.hosts.includes(checked.hostname)) return '';
        return 'pub-'+site.key+':'+encodeURIComponent(checked.href);
      }
      if (u.hostname === 'book.douban.com') {
        const sid = (u.pathname.match(/^\/subject\/(\d+)\/?$/) || [])[1];
        return sid ? 'douban-' + sid : '';
      }
      return u.hostname === 'www.books.com.tw' ? (u.pathname.match(/^\/products\/([A-Za-z0-9_-]+)\/?$/) || [])[1] || '' : '';
    } catch { return ''; }
  }
  function originalCover(raw) {
    try {
      let u = new URL(String(raw || '').replace(/&amp;/g, '&'), typeof location!=='undefined' && location.protocol==='https:' ? location.href : 'https://www.books.com.tw');
      if (u.username || u.password || u.port || !['http:','https:'].includes(u.protocol)) return '';
      if(SITE_ADAPTERS.some(s=>s.coverHosts.includes(u.hostname))) {
        // Publisher/CDN URLs may carry required signatures and have no extension.
        if(u.protocol!=='https:' || u.pathname==='/' || /(?:login|captcha)/i.test(u.pathname)) return '';
        const taazeImage=u.hostname==='media.taaze.tw' && /^\/show(?:ProdImage|Thumbnail|LargeImage)\.html$/.test(u.pathname) && /^111[0-9]{8}$/.test(u.searchParams.get('sc') || '');
        if(u.hostname==='media.taaze.tw' && !taazeImage) return '';
        const noExtension=taazeImage || u.hostname==='img.91app.com' && /^\/webapi\/imagesV3\//.test(u.pathname) || u.hostname==='www.shogakukan.co.jp' && /^\/thumbnail\/books\/\d+\/?$/.test(u.pathname);
        if(!/\.(jpg|jpeg|png|webp|gif)$/i.test(u.pathname) && !noExtension) return '';
        u.hash='';return u.href;
      }
      if (['img.doubanio.com','img1.doubanio.com','img2.doubanio.com','img3.doubanio.com','img9.doubanio.com'].includes(u.hostname)) {
        if (!/^\/view\/subject\/[a-z]+\/public\/[A-Za-z0-9_-]+\.(jpg|jpeg|png|webp)$/i.test(u.pathname)) return '';
        u.protocol='https:'; u.hash=''; u.search=''; return u.href;
      }
      if (u.pathname === '/image/getImage') u = new URL(u.searchParams.get('i'), u);
      if (!['im1.book.com.tw', 'im2.book.com.tw', 'www.books.com.tw'].includes(u.hostname) ||
          !u.pathname.includes('/img/') || !/\.(jpg|jpeg|png|webp|gif)$/i.test(u.pathname)) return '';
      u.protocol = 'https:'; u.search = ''; u.hash = '';
      return u.href;
    } catch { return ''; }
  }
  function isbn13(raw) {
    let s = String(raw || '').replace(/[\s-]/g, '').toUpperCase();
    if (/^\d{9}[\dX]$/.test(s)) {
      if ([...s].reduce((sum, c, i) => sum + (c === 'X' ? 10 : +c) * (10 - i), 0) % 11) return '';
      s = '978' + s.slice(0, 9);
      s += (10 - [...s].reduce((sum, c, i) => sum + +c * (i % 2 ? 3 : 1), 0) % 10) % 10;
    }
    return /^(978|979)\d{10}$/.test(s) &&
      [...s].reduce((sum, c, i) => sum + +c * (i % 2 ? 3 : 1), 0) % 10 === 0 ? s : '';
  }
  function numberText(raw) {
    let s = String(raw).normalize('NFKC').replace(/[壹貳贰參叁肆伍陸陆柒捌玖拾佰兩两]/g,
      c => ({壹:'一',貳:'二',贰:'二',參:'三',叁:'三',肆:'四',伍:'五',陸:'六',陆:'六',柒:'七',捌:'八',玖:'九',拾:'十',佰:'百',兩:'二',两:'二'})[c]);
    if (/^\d+(\.\d+)?$/.test(s)) return String(Number(s));
    if (!/^[零〇一二三四五六七八九十百]+$/.test(s)) return '';
    const digits = {零:0,〇:0,一:1,二:2,三:3,四:4,五:5,六:6,七:7,八:8,九:9};
    if (!/[十百]/.test(s)) return String(Number([...s].map(c => digits[c]).join('')));
    let sum = 0, n = 0;
    for (const c of s) { if (c === '十' || c === '百') { sum += (n || 1) * (c === '十' ? 10 : 100); n = 0; } else n = digits[c]; }
    return String(sum + n);
  }
  const stripEbookSuffix = title => clean(title).replace(/\s*(?:[（(【[]\s*)?(?:電子書|电子书|e-?book)\s*(?:[）)】\]])?\s*$/i, '').trim();
  function namedPart(title) {
    for(const pattern of [/[([【]\s*([上中下])\s*(?:巻|卷|冊|册|部)?\s*[)\]】]/g, /\s+([上中下])\s*(?:巻|卷|冊|册|部)?(?=\s|[:：~～〜—–-]|$)/g]) {
      for(const match of title.matchAll(pattern)) {
        let root=title.slice(0,match.index).trim();
        if(!root || (root.match(/[([【]/g) || []).length!==(root.match(/[)\]】]/g) || []).length || /(?:\s|[(/、+])(?:上|中)(?:\s*[/、+])?$/.test(root)) continue;
        const suffix=title.slice(match.index+match[0].length).trim();
        const badge=suffix.match(/^[([【]?\s*(無修正(?:版)?|无修正(?:版)?|漫畫(?:版)?|漫画(?:版)?|小說|小说|小説)\s*[)\]】]?$/);
        if(badge) root+=' '+badge[1];
        return {root,number:match[1],index:match.index};
      }
    }
    return null;
  }
  function identity(title) {
    let t = clean(title).normalize('NFKC').replace(/^博客來[-－—:：]\s*/, '');
    const boxed = /套書|套书|合售|盒裝|盒装|全套|\d+\s*[-~～]\s*\d+\s*[集卷冊册]/.test(t);
    const ebook = /電子書|电子书|e-?book/i.test(t);
    if (ebook) t = stripEbookSuffix(t);
    let kind = /公式|導讀|导读|設定集|设定集|資料集|资料集/.test(t) ? 'guidebook' :
      /畫集|画集|插畫集|插画集/.test(t) ? 'artbook' :
      /外傳|外传|番外/.test(t) ? 'extra' : 'main';
    const editions = [[/豪華典藏版|豪华典藏版/, 'deluxe_collection'], [/完全版/, 'kanzenban'],
      [/愛藏版|爱藏版/, 'aizouban'], [/典藏版/, 'deluxe_collection'], [/珍藏版/, 'collector'], [/新裝(?:版)?|新装(?:版)?/, 'new_edition'],
      [/文庫版|文库版/, 'bunko'], [/全彩版/, 'full_color'], [/豪華版|豪华版/, 'luxury']];
    const edition = editions.find(([re]) => re.test(t));
    const editionType = edition?.[1] || 'tankobon';
    const editionLabel = edition ? (t.match(edition[0]) || [])[0] : '单行本';
    const print = /首刷限定/.test(t) ? 'first_print_limited' : /首刷/.test(t) ? 'first_print' : /限定/.test(t) ? 'limited' : 'standard';
    t = t.replace(/[（(【[][^）)】\]]*(?:首刷|限定)[^）)】\]]*[）)】\]]/g, '').trim();
    // Editions occur before OR after the volume. Remove them before parsing,
    // but retain their separate identity so different editions never collapse.
    if (edition) t = t.replace(edition[0], '').replace(/\(\s*\)/g, '').trim();
    t = t.replace(/^[‧·・：:、\-－—\s]+/, '').replace(/[‧·・：:、\-－—\s]+$/, '').trim();
    t = t.replace(/[【[]([^】\]]+)[】\]]/g, '($1)');
    const completed = /(?:[(\s](?:完[結结]?|END)[)]?|[\d一二三四五六七八九十]+[)]?完[)]?)\s*$/i.test(t);
    t = t.replace(/[(]\s*(?:完[結结]?|END)\s*[)]\s*$|\s+(?:完[結结]?|END)\s*$/i, '').trim();
    let volume = '', sequence = 'numeric', count = null;
    const single = t.match(/(?:[(\s]|^)(全[一1](?:[卷冊册集])?|一本完|單冊完結|单册完结|全)[)]?\s*$/);
    const position = t.match(/(?:[(\s](上|中|下)(?:卷|冊|册)?|(?:卷|冊|册)(上|中|下)|(上|中|下)(?:卷|冊|册))[)]?\s*$/);
    const digits='[\\d零〇一二三四五六七八九十百壹貳贰參叁肆伍陸陆柒捌玖拾佰兩两]+(?:\\.\\d+)?';
    const numeric = t.match(new RegExp('(?:第\\s*|[(\\s])('+digits+')\\s*(?:卷|巻|集|冊|册)?\\s*(?:[)]?\\s*完)?[)]?\\s*$')) ||
      t.match(new RegExp('(?:第\\s*|[(\\s])('+digits+')\\s*(?:卷|巻|集|冊|册)?[)]?\\s*[:：]\\s*.+$')) ||
      t.match(new RegExp('(?:卷|巻|集|冊|册)\\s*('+digits+')[)]?\\s*$')) ||
      t.match(new RegExp('('+digits+')\\s*(?:卷|巻|集|冊|册)[)]?\\s*$')) || t.match(/(?<=[\u3400-\u9fff])([0-9]{1,3})\s*$/);
    const named = !boxed ? namedPart(t) : null;
    const suffix = single || position || numeric;
    if (single) { volume = '全'; sequence = 'single'; count = 1; }
    else if (position) { volume = position[1] || position[2] || position[3]; }
    else if (numeric) volume = numberText(numeric[1]);
    if (named && (!numeric || named.index<=numeric.index)) { t=named.root; volume=named.number; }
    else if (suffix) t = t.slice(0, suffix.index).replace(/\(\s*$/,'').trim();
    t = t.replace(/[\s:：\-－—]+$/, '').trim();
    if (!volume && kind !== 'main') volume = KINDS.find(([k]) => k === kind)[1].replace(' / ', '');
    return {work_title:t || title, volume_number:volume, volume_sequence_format:sequence,
      published_volume_count:count, serial_status:single || completed ? 'completed' : 'unknown',
      edition_type:editionType, edition_label:editionLabel, print_variant:print, volume_kind:kind,
      content_type:['guidebook', 'artbook'].includes(kind) ? kind : /小說|小说/.test(title) ? 'novel' : 'manga',
      sequence_role:kind === 'main' ? 'primary' : 'supplement', is_boxed:boxed, ebook};
  }
  function enrichIdentity(parts, doc, summary) {
    const categories=[...doc.querySelectorAll('#db-tags-section,.tags-body,.breadcrumb,#breadcrumb,.type04_breadcrumb')].map(n=>clean(n.textContent)).join(' ');
    if(parts.volume_kind==='main' && /小說|小说/.test(categories) && !/漫畫|漫画/.test(categories)) parts.content_type='novel';
    if(!parts.volume_number && !parts.is_boxed && /(?:本書|本书|本作品)\s*(?:[，,:：]\s*)?(?:(?:為|为|是|共)\s*)?(?:全[一1][冊册卷]|單冊完結|单册完结)/.test(summary || ''))
      Object.assign(parts,{volume_number:'全',volume_sequence_format:'single',published_volume_count:1,serial_status:'completed'});
    return parts;
  }
  const workKey = value => clean(value).normalize('NFKC').replace(/\s+/g,'').toLowerCase();
  // Explicit desktop identities (backend/app/catalog/normalization.py). Do not
  // infer Kadokawa's region from a partial name or merge Taiwan with Tianwen.
  const PUBLISHERS = {
    '东立':['東立','东立出版','東立出版','东立出版社','東立出版社','台湾东立','台灣東立','臺灣東立'],
    '尖端':['尖端出版','尖端出版社','尖端漫画','尖端漫畫','台湾尖端','台灣尖端','臺灣尖端'],
    '青文':['青文出版','青文出版社','台湾青文','台灣青文','臺灣青文'],
    '长鸿':['長鴻','长鸿出版','長鴻出版','长鸿出版社','長鴻出版社','台湾长鸿','台灣長鴻'],
    '台湾角川':['台灣角川','臺灣角川','台湾角川出版','台灣角川出版','台湾角川出版社','台灣角川出版社'],
    '东贩':['台湾东贩','台灣東販','臺灣東販','台湾东贩出版','台灣東販出版','東販'],
    '大然':['大然文化','大然出版社'],'时报文化':['時報文化','时报出版','時報出版'],
    '盖亚文化':['蓋亞文化','盖亚','蓋亞'],'枫书坊':['楓書坊','枫书坊文化','楓書坊文化'],
    '远流':['遠流','远流出版','遠流出版'],'联经':['聯經','联经出版','聯經出版'],'脸谱':['臉譜','脸谱出版','臉譜出版'],
    '玉皇朝':['玉皇朝出版','玉皇朝出版集团','玉皇朝出版集團'],'天下出版':['天下出版社'],
    '文化传信':['文化傳信'],'正文社':['正文社出版'],'香港青文':['香港青文出版','香港青文出版社'],'海洋创作':['海洋創作'],
    '天闻角川':['天聞角川','广州天闻角川动漫','廣州天聞角川動漫','广州天闻角川','廣州天聞角川','湖南天闻动漫传媒','湖南天聞動漫傳媒','湖南天闻动漫','湖南天聞動漫'],
    '人民邮电出版社':['人民郵電出版社','人邮','人郵'],'中信出版社':['中信出版集团','中信出版集團'],'新星出版社':[],
    '浙江人民美术出版社':['浙江人民美術出版社'],'上海人民美术出版社':['上海人民美術出版社'],
    '四川美术出版社':['四川美術出版社'],'湖南美术出版社':['湖南美術出版社'],
    '北京联合出版公司':['北京聯合出版公司'],'接力出版社':[],'中国少年儿童出版社':['中國少年兒童出版社'],'二十一世纪出版社':['二十一世紀出版社'],
    '集英社':['株式会社集英社'],'讲谈社':['講談社','株式会社讲谈社','株式会社講談社'],
    '小学馆':['小學館','小学館','株式会社小学馆','株式会社小学館'],'KADOKAWA':['株式会社KADOKAWA','日本角川'],
    '白泉社':['株式会社白泉社'],'秋田书店':['秋田書店','株式会社秋田书店','株式会社秋田書店'],
    '芳文社':['株式会社芳文社'],'双叶社':['雙葉社','双葉社','株式会社双葉社'],'新潮社':['株式会社新潮社'],
    '少年画报社':['少年畫報社','少年画報社'],'SQUARE ENIX':['スクウェア・エニックス'],
  };
  const publisherAliases = new Map(Object.entries(PUBLISHERS).flatMap(([name,aliases])=>[name,...aliases].map(alias=>[workKey(alias),name])));
  function canonicalPublisher(value) {
    const raw=clean(value).normalize('NFKC').replace(/\s*(?:訂閱出版社新書快訊|订阅出版社新书快讯|追蹤出版社|追踪出版社|新功能介紹|新功能介绍).*$/,'')
      .replace(/^(?:臺?北|台北|臺?中|台中|高雄|新北|桃園|桃园|香港|九龍|九龙|上海|北京|广州|廣州)市?\s*[:：]/,'').trim();
    const stem=raw.replace(/(?:出版社)?(?:文化|出版|图书|圖書|动漫|動漫|传媒|傳媒|事业|事業|企业|企業)?(?:集团|集團)?(?:股份)?(?:有限责任|有限責任)?有限公司$/,'');
    return publisherAliases.get(workKey(raw)) || publisherAliases.get(workKey(stem)) || stem;
  }
  const publisherKey = value => workKey(canonicalPublisher(value));
  function relatedSeries(doc, base) {
    const source=SITE_ADAPTERS.find(s=>s.hosts.includes(new URL(base).hostname));
    if(source) {
      const raw=source.series?.(doc,base);
      try {const u=new URL(raw,base);return raw && u.protocol==='https:' && !u.username && !u.password && !u.port && source.hosts.includes(u.hostname) && source.list(u) ? [u.href] : [];} catch {return [];}
    }
    const douban=new URL(base).hostname==='book.douban.com';
    const anchors=[...doc.querySelectorAll(douban ? '#info a[href]' : 'a[href]')];
    return [...new Set(anchors.filter(a=>!a.closest('#bcd-panel,.flash_pic,header,nav,.banner')) .map(a=>{
      try {
        const u=new URL(a.getAttribute('href'),base);
        if(u.protocol!=='https:' || u.username || u.password || u.port) return '';
        if(douban) return u.hostname==='book.douban.com' && /^\/series\/\d+\/?$/.test(u.pathname) ? u.href : '';
        // loc=series is the site's explicit product-series relationship, not a
        // top banner or a loosely related marketing collection.
        return u.hostname==='activity.books.com.tw' && /^\/crosscat\/show\/[A-Za-z0-9]+\/?$/.test(u.pathname) && u.searchParams.get('loc')==='series' ? u.href : '';
      } catch { return ''; }
    }).filter(Boolean))];
  }
  function supplementIds(text, douban=false, siteKey=CURRENT_SITE?.key || '') {
    const links=String(text).trim().split(/[\s,，;；]+/).filter(Boolean);
    if(!links.length || links.length>MAX_ITEMS) throw failure('请粘贴 1–300 条本站单本商品链接，每行一条。');
    const ids=links.map(url=>/^https:\/\//i.test(url) ? productId(url) : '');
    if(ids.some(id=>!id || (siteKey ? publisherRef(id)?.site.key!==siteKey : id.startsWith('pub-') || id.startsWith('douban-')!==douban))) throw failure('只接受本站 HTTPS 单本商品链接；不要粘贴系列、广告或其他网站的链接。');
    return [...new Set(ids)];
  }
  function directoryGaps(items) {
    const groups=new Map();
    for(const item of items) {
      const e=item.record?.entry || identity(item.title);
      if(e.is_boxed || e.sequence_role!=='primary' || !/^\d+$/.test(e.volume_number)) continue;
      const key=JSON.stringify([workKey(e.work_title),e.edition_type]);
      if(!groups.has(key)) groups.set(key,{title:e.work_title,numbers:new Set()});
      groups.get(key).numbers.add(+e.volume_number);
    }
    return [...groups.values()].flatMap(({title,numbers})=>{
      const max=Math.max(...numbers);
      if(numbers.size<2 || max>MAX_ITEMS) return [];
      const missing=Array.from({length:max},(_,i)=>i+1).filter(n=>!numbers.has(n));
      return missing.length ? [`${title}：当前目录未包含第 ${missing.join('、')} 卷（不是书库缺卷判断）`] : [];
    });
  }
  const setKey = e => JSON.stringify([workKey(e.work_title),e.region,publisherKey(e.publisher),e.edition_type,e.content_type,e.volume_kind || 'main']);
  function prepareSets(records, pageTitle='', forcedTitle='') {
    const hint=identity(pageTitle), hintKey=workKey(hint.work_title);
    const rows=records.map(rec=>({ ...rec, entry:{...rec.entry,publisher:canonicalPublisher(rec.entry.publisher) || null} }));
    for(const rec of rows) {
      const e=rec.entry;
      if(rec.blocked || e.is_boxed) continue;
      if(forcedTitle) e.work_title=clean(forcedTitle);
      // A series heading may clean punctuation/spacing, but is not proof that
      // every book on a publisher's collection page belongs to one work.
      else if(hintKey && workKey(e.work_title)===hintKey) e.work_title=hint.work_title;
    }
    const baseKey=e=>JSON.stringify([workKey(e.work_title),e.region,e.edition_type,e.content_type,e.volume_kind || 'main']);
    const buckets=new Map();
    for(const rec of rows.filter(r=>!r.blocked)) {
      const key=baseKey(rec.entry); if(!buckets.has(key)) buckets.set(key,[]); buckets.get(key).push(rec.entry);
    }
    for(const entries of buckets.values()) {
      const publishers=new Map(entries.filter(e=>clean(e.publisher)).map(e=>[publisherKey(e.publisher),e.publisher]));
      // Missing publisher can join exactly one known publisher, never bridge
      // two conflicting publishers. Canonical spelling must reach the export.
      for(const e of entries) {
        if(publishers.size===1) e.publisher=[...publishers.values()][0];
        e.work_title=entries[0].work_title;
      }
    }
    return rows;
  }
  function inferSets(records, directory=[], complete=false) {
    records=prepareSets(records);
    const knownDirectory=prepareSets(directory.map(i=>i.record).filter(Boolean));
    let offset=0;
    directory=directory.map(i=>i.record ? {...i,record:knownDirectory[offset++]} : i);
    const groups=new Map();
    for(const rec of records.filter(r=>!r.blocked)) {
      const key=setKey(rec.entry); if(!groups.has(key)) groups.set(key,[]); groups.get(key).push(rec);
    }
    const results=new Map();
    for(const [key,rows] of groups) {
      const first=rows[0].entry, primary=rows.filter(r=>r.entry.sequence_role==='primary');
      const matchesWork=e=>workKey(e.work_title)===workKey(first.work_title) && e.edition_type===first.edition_type;
      const known=[...records,...directory.map(item=>item.record).filter(Boolean)].filter(r=>!r.blocked && matchesWork(r.entry));
      const ambiguous=new Set(known.map(r=>setKey(r.entry))).size>1;
      const evidence=directory.map(item=>item.record ? (!item.record.blocked && setKey(item.record.entry)===key ? item.record.entry : null) : ambiguous ? null : identity(item.title))
        .filter(e=>e && !e.is_boxed && e.sequence_role==='primary' && matchesWork(e));
      const entries=[...evidence,...primary.map(r=>r.entry)], numbers=entries.map(e=>e.volume_number).filter(Boolean);
      const numeric=numbers.filter(n=>/^\d+$/.test(n)).map(Number);
      const terminals=[...new Set(entries.filter(e=>e.serial_status==='completed' && /^\d+$/.test(e.volume_number)).map(e=>+e.volume_number))];
      const defaults={}, notes=[];
      const same = (field, empty) => {const values=rows.map(r=>r.entry[field]).filter(v=>v != null && v!=='' && (!Array.isArray(v) || v.length)); const unique=[...new Map(values.map(v=>[JSON.stringify(v),v])).values()]; return unique.length===1 ? unique[0] : empty;};
      Object.assign(defaults,{work_title:first.work_title,region:first.region,publisher:same('publisher',''),authors:same('authors',[]),edition_type:first.edition_type,content_type:first.content_type});
      if(new Set(rows.filter(r=>r.entry.authors?.length).map(r=>JSON.stringify(r.entry.authors))).size>1) notes.push('作者信息不一致，保留各卷作者，请核对后再统一');
      if(primary.length && numbers.length && numbers.every(n=>n==='全')) Object.assign(defaults,{volume_sequence_format:'single',published_volume_count:1,serial_status:'completed'});
      else if(primary.length && complete && numbers.includes('上') && numbers.includes('下') && numbers.every(n=>['上','中','下'].includes(n))) {
        const middle=numbers.includes('中'); Object.assign(defaults,{volume_sequence_format:middle?'upper_middle_lower':'upper_lower',published_volume_count:middle?3:2,serial_status:'completed'});
      } else if(terminals.length===1 && numeric.length===numbers.length && Math.max(...numeric)<=terminals[0]) Object.assign(defaults,{volume_sequence_format:'numeric',published_volume_count:terminals[0],serial_status:'completed'});
      else {
        if(terminals.length) notes.push('完结卷号与其他正篇冲突，请确认总卷数');
        if(numbers.some(n=>['上','中','下'].includes(n))) notes.push('上下册目录尚未确认完整，请选择上下或上中下');
        else if(primary.length) defaults.volume_sequence_format='numeric';
      }
      if(!defaults.published_volume_count) notes.push('总卷数待确认，不按网站商品条目数推算');
      results.set(key,{defaults,notes});
    }
    return results;
  }
  function seriesUrl(value, base, seriesId) {
    try {
      const u=new URL(value,base);
      if(u.origin!=='https://book.douban.com' || u.username || u.password || u.port || u.pathname.replace(/\/$/,'')!=='/series/'+seriesId) return '';
      const page=u.searchParams.get('page') || '1', order=u.searchParams.get('order') || new URL(base).searchParams.get('order') || 'count';
      if(!/^[1-9]\d*$/.test(page) || +page>30 || !['count','time'].includes(order)) return '';
      return `https://book.douban.com/series/${seriesId}?order=${order}&page=${+page}`;
    } catch { return ''; }
  }
  function seriesPage(doc, url, seriesId) {
    const canonical=seriesUrl(url,url,seriesId);
    if(!canonical || /异常请求|安全验证|驗證|验证/.test(doc.title) || doc.querySelector('form[action*="captcha"],input[name*="captcha"],form[action*="passport/login"]')) throw failure('目录需要人工验证或返回其他页面；未合并，请处理后继续。','challenge');
    if(!doc.querySelector('#content .article .subject-list')) throw failure('丛书目录未加载完整，未标记为读取完成。');
    const order=new URL(canonical).searchParams.get('order');
    const anchors=[...doc.querySelectorAll('.paginator a[href]')];
    for(const anchor of anchors) {
      try {
        const u=new URL(anchor.getAttribute('href'),canonical);
        if(u.origin==='https://book.douban.com' && u.pathname.replace(/\/$/,'')==='/series/'+seriesId && +(u.searchParams.get('page') || 1)>30) throw failure('目录超过 30 页，请缩小范围；不能标记为完整。','limit');
      } catch(err) {if(err.kind==='limit') throw err;}
    }
    const links=anchors.map(a=>seriesUrl(a.getAttribute('href'),canonical,seriesId)).filter(u=>u && new URL(u).searchParams.get('order')===order);
    const first=new URL(canonical); first.searchParams.set('page','1');
    const items=collectItems(doc);
    if(!items.length) throw failure('丛书目录没有有效书目，未标记为读取完成。');
    return {url:canonical,links:[...new Set([first.href,...links])],items};
  }
  // Same five footprints and distance rule as backend/app/catalog/trim_format.py.
  function trimFormat(spec) {
    const explicit = spec.match(/\b(64|36|32|25|16)\s*[kKＫ開开]/);
    if (explicit) return explicit[1] + 'K';
    const m = spec.match(/(\d{1,3}(?:\.\d+)?)\s*[xX×＊*]\s*(\d{1,3}(?:\.\d+)?)(?:\s*[xX×＊*]\s*(\d{1,3}(?:\.\d+)?))?\s*(mm|cm|毫米|公分)?/);
    if (!m) return null;
    const sides = m.slice(1, 4).filter(Boolean).map(v => +v / (/mm|毫米/i.test(m[4]) ? 10 : 1)).sort((a,b) => b-a);
    const [h,w] = sides;
    if (!(w >= 8 && w <= 25 && h >= 12 && h <= 35)) return null;
    return [[10.5,14.8,'64K'], [11.3,17.2,'36K'], [11.5,18.9,'36K'], [13,19,'32K'], [14.8,21,'25K'], [19,26,'16K']]
      .sort((a,b) => ((w-a[0])/a[0])**2 + ((h-a[1])/a[1])**2 - ((w-b[0])/b[0])**2 - ((h-b[1])/b[1])**2)[0][2];
  }
  function description(doc) {
    for (const heading of doc.querySelectorAll('h2,h3,h4,h5,h6,[role="heading"]')) {
      if (!/^(內容簡介|内容简介|內容简介|内容簡介)$/.test(clean(heading.textContent).replace(/[\s*＊★☆※:：【】[\]]/g, ''))) continue;
      let parent = heading.parentElement;
      for (let depth = 0; depth < 3 && parent && !/^(BODY|HTML)$/.test(parent.tagName); depth++, parent = parent.parentElement) {
        const ordered = [...parent.querySelectorAll('*')], start = ordered.indexOf(heading);
        const body = [...parent.querySelectorAll('.bd .content')].find(node => {
          const end = ordered.indexOf(node);
          return end > start && !ordered.slice(start+1, end).some(n => /^H[1-6]$/.test(n.tagName) || n.getAttribute('role') === 'heading');
        });
        if (body) {
          const copy = body.cloneNode(true);
          copy.querySelectorAll('script,style').forEach(n => n.remove());
          copy.querySelectorAll('br').forEach(n => n.replaceWith('\n'));
          copy.querySelectorAll('p,div,li').forEach(n => { n.prepend('\n'); n.append('\n'); });
          return copy.textContent.split('\n').map(clean).filter(Boolean).join('\n').slice(0, 20000);
        }
        if (parent.querySelectorAll('h2,h3,h4,h5,h6,[role="heading"]').length > 1) break;
      }
    }
    return '';
  }
  function failure(message, kind = 'error', delay = 0) { return Object.assign(new Error(message), {kind, delay}); }
  function parseBook(html, id, finalUrl = productUrl(id)) {
    if(id.startsWith('pub-')) return parsePublisher(html,id,finalUrl);
    if (id.startsWith('douban-')) return parseDouban(html,id,finalUrl);
    if (/連線暫時異常|Connection is temporarily unavailable|_cf_chl_opt|cf-chl-widget|Just a moment\.\.\./i.test(html))
      throw failure('博客来要求验证或暂时限制访问。打开该商品处理后再继续。', 'challenge', 60000);
    if (/\/member\/login/.test(finalUrl) || /本商品為限制級/.test(html) && !/class=["'][^"']*type02_p003/.test(html))
      throw failure('需要登录或年龄确认。打开该商品完成后，回到这里继续。', 'login');
    if (productId(finalUrl) !== id) throw failure('返回了其他页面，未保存。请打开商品核对。');
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const canonical = doc.querySelector('link[rel="canonical"]')?.getAttribute('href') || doc.querySelector('meta[property="og:url"]')?.content;
    if (canonical && productId(canonical) !== id) throw failure('商品编号与返回页面不一致，未保存。');
    const title = clean(doc.querySelector('h1')?.textContent);
    const info = doc.querySelector('.type02_p003');
    // Hidden follow-author dialogs precede the real author in the same <li>.
    // Remove UI controls from a clone, not from the customer's live page.
    const metadata=(info || doc.body).cloneNode(true);
    metadata.querySelectorAll('.trace_box,.type02_btn02,.trace_txt,.help,script,style,[hidden],#bcd-panel').forEach(n=>n.remove());
    const metadataFields=[...metadata.querySelectorAll('li,p,span')].map(n=>clean(n.textContent));
    const fieldNodes=[...metadataFields,...[...doc.querySelectorAll('li,p,span')].filter(n=>!n.closest('#bcd-panel,.type02_p003')).map(n=>clean(n.textContent))];
    function field(label) {
      const re = new RegExp('^' + label + '\\s*[：:]\\s*(.+?)(?=\\s+(?:作者|出版社|出版日期|ISBN|EAN|條碼|定價|優惠價|規格)\\s*[：:]|$)');
      const nodes=['作者','出版社'].includes(label) ? metadataFields : fieldNodes;
      return nodes.map(t => (t.match(re) || [])[1]).find(Boolean) || '';
    }
    if (!title || !info && !field('ISBN') && !field('規格')) throw failure('页面资料未加载完整，未缓存；请打开商品后重试。');
    const intro=description(doc) || null;
    const parts = enrichIdentity(identity(title),doc,intro);
    const ld = [...doc.querySelectorAll('script[type="application/ld+json"]')].flatMap(node => {
      try { const data = JSON.parse(node.textContent); return Array.isArray(data) ? data : data['@graph'] || [data]; } catch { return []; }
    }).find(obj => obj && /Book|Product/.test(String(obj['@type'])) && (!obj.url || productId(obj.url) === id)) || {};
    const rawCode = clean(field('ISBN') || field('EAN') || field('條碼') || field('条码') || ld.isbn);
    const isbn = isbn13(rawCode);
    const oldCode = !isbn && /^\d{10,13}$/.test(rawCode.replace(/[\s-]/g, '')) && !/^(978|979)/.test(rawCode) ? rawCode.replace(/[\s-]/g, '') : null;
    const strip = text => clean(text).replace(/\s*(?:追蹤作者|追踪作者|新功能介紹|訂閱新書快訊|订阅新书快讯|相關資訊|相關介紹|[[【]修改[\]】]).*$/, '').trim();
    const author = strip(field('作者'));
    const spec = field('規格') || field('规格');
    const rate = spec + ' ' + field('分級');
    const price = field('定價').match(/(?:NT\$\s*)?([\d,]+(?:\.\d+)?)\s*元/);
    const date = field('出版日期').match(/(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
    const pubDate = date ? `${date[1]}-${date[2].padStart(2,'0')}-${date[3].padStart(2,'0')}` : null;
    const cover = [doc.querySelector('meta[property="og:image"]')?.content, ...[...doc.querySelectorAll('img.cover')].map(n => n.getAttribute('data-src') || n.src)].map(originalCover).find(Boolean) || '';
    const {ebook, ...identityFields} = parts;
    const volumeTitle = ebook ? stripEbookSuffix(title) : title;
    const entry = {...identityFields, volume_title:volumeTitle.slice(0,240), work_title:parts.work_title.slice(0,240),
      region:'TW', authors:author ? author.split(/[、,/／;&＆]+/).map(clean).filter(Boolean).slice(0,30) : [],
      publisher:canonicalPublisher(strip(field('出版社'))).slice(0,160) || null, isbn:isbn || null, legacy_barcode:oldCode,
      pub_date:pubDate, description:intro, format:parts.content_type === 'manga' && !parts.is_boxed ? trimFormat(spec) : null,
      audience_rating:/限制[級级]|18\s*禁/.test(rate) ? 'restricted' : /輔導[級级]|辅导[級级]/.test(rate) ? 'guidance' : /普通[級级]/.test(rate) ? 'general' : 'pending',
      list_price:price ? Number(price[1].replace(/,/g,'')) : null, currency:price ? 'TWD' : null,
      notes:'资料来源：' + productUrl(id) + (ebook ? '（博客来无实体商品时采用电子书资料；缺失资料请人工补录）' : '')};
    const warnings = [];
    if (!isbn) warnings.push(oldCode ? '仅旧条码：不能按 ISBN 自动搜书' : rawCode ? '书号校验失败：请核对' : 'ISBN 暂缺');
    if (!entry.description) warnings.push('简介暂缺');
    if (!entry.format) warnings.push('开本暂缺');
    if (!cover) warnings.push('封面暂缺');
    if (ebook) warnings.unshift('博客来当前仅提供电子书资料；缺失项目请人工补录');
    const supportedId=/^001\d+$/.test(id) || ebook && /^E\d+$/i.test(id);
    return {id, title, entry, cover, warnings, ebook, metadataRevision:1, blocked:!supportedId ? '仅支持中文博客来商品' : parts.is_boxed ? '盒装 / 多册合售请在书库按套装核对，不能当作一卷导出' : '', fetchedAt:Date.now()};
  }
  function regionFromIsbn(code) {
    return /^(978626|978957|978986)/.test(code) ? 'TW' : /^(978962|978988)/.test(code) ? 'HK' : /^9787/.test(code) ? 'CN' : /^9784/.test(code) ? 'JP' : 'UNKNOWN';
  }
  function citiBaseUrl(value) {
    const u=new URL(value);u.searchParams.delete('bcd_variant');return u.href;
  }
  function citiProductData(doc) {
    for(const script of doc.querySelectorAll('script:not([src])')) {
      const match=script.textContent.match(/app\.value\(['"]product['"],\s*(\{[\s\S]*?\})\s*\);/);
      if(!match) continue;
      try {const data=JSON.parse(match[1]);if(Array.isArray(data.field_titles) && Array.isArray(data.variations))return data;}
      catch { /* Read the public JSON value only; never execute page scripts. */ }
    }
    return null;
  }
  function citiVariants(doc,value) {
    const data=citiProductData(doc);if(!data)return [];
    const names=data.field_titles.map(v=>clean(v.name_translations?.['zh-hant']));
    const bookIndex=names.indexOf('書目名稱'), volumeIndex=names.indexOf('期數');
    if(volumeIndex<0 || names.length!== (bookIndex<0 ? 1 : 2)) return [];
    const base=citiBaseUrl(value), seen=new Set(), result=[];
    for(const variant of data.variations) {
      const fields=variant.fields_translations?.['zh-hant'];
      if(!/^[a-f0-9]{24}$/.test(variant.key) || seen.has(variant.key) || !Array.isArray(fields))continue;
      const book=clean(bookIndex<0 ? data.title_translations?.['zh-hant'] : fields[bookIndex]), volume=clean(fields[volumeIndex]);
      if(!book || !volume)continue;
      const url=new URL(base);url.searchParams.set('bcd_variant',variant.key);url.hash='';
      seen.add(variant.key);result.push({url:url.href,title:book+' '+volume,book,volume,variant});
    }
    return result;
  }
  function citiVariantRaw(doc,value) {
    const row=citiVariants(doc,value).find(v=>v.variant.key===new URL(value).searchParams.get('bcd_variant'));
    if(!row)throw failure('本页未提供所选书目／期数的公开数据，请核对原站选项。');
    const position=row.volume.match(/^(上|中|下)(?:卷|冊)?$/), numeric=row.volume.match(/^(?:第)?(\d+)(?:期|卷|冊)$/), single=/^全[一1](?:冊|卷)$/.test(row.volume);
    const volume=position?.[1] || (numeric ? String(Number(numeric[1])) : single ? '全' : '');
    const hard=/硬皮/.test(row.book), collector=/珍藏本/.test(row.book);
    return {title:row.title,work_title:row.book,volume_number:volume,
      volume_sequence_format:position ? 'upper_lower' : single ? 'single' : 'numeric',
      publisher:'文化傳信',authors:[],isbn:'',description:'',
      edition_type:collector?'collector':'tankobon',edition_label:collector ? '珍藏本'+(hard?'／限量硬皮版':'') : identity(row.book).edition_label,
      print_variant:/限量/.test(row.book)?'limited':'standard',
      cover:row.variant.media?.images?.original?.url || '',list_price:row.variant.price?.dollars ?? null,currency:row.variant.price?.currency_iso || '',
      metadata_note:'同页书目：'+row.book+'；期数：'+row.volume+'；原站变体 ID：'+row.variant.key+'。未将整页简介或其他书目的作者、书号套用到本册。',
      blocked:volume?'':'原站变体册号尚未识别，请核对后补录'};
  }
  function parsePublisher(html,id,finalUrl=productUrl(id),apiData=null,publicSearchCode=null) {
    const ref=publisherRef(id);
    const variant=ref?.site.key==='citicomics' && new URL(ref.url).searchParams.has('bcd_variant');
    const returned=variant ? productId(citiBaseUrl(finalUrl)) : productId(finalUrl), expected=variant ? productId(citiBaseUrl(ref.url)) : id;
    if(!ref || returned!==expected) throw failure('出版社返回了其他页面，未保存；请打开商品核对。');
    const doc=new DOMParser().parseFromString(html,'text/html');
    if(/Just a moment|Attention Required|安全验证|安全驗證/i.test(doc.title) || doc.querySelector('form[action*="login"],input[name*="captcha"],#challenge-form'))
      throw failure('出版社需要登录或人工验证，请在本站完成后继续。','challenge');
    const canonical=doc.querySelector('link[rel="canonical"]')?.href;
    if(canonical && productId(new URL(canonical,finalUrl).href)!==expected) throw failure('出版社商品编号与页面不一致，未保存。');
    const raw=apiData && ref.site.parseApi ? ref.site.parseApi(apiData) : ref.site.parse(doc,ref.url,publicSearchCode);
    if(!raw || !clean(raw.title)) throw failure('出版社详情未加载完整，未缓存。请在商品页面等待加载后重试。');
    const title=clean(raw.title), parts=enrichIdentity(identity(title),doc,raw.description), {ebook,...fields}=parts;
    const code=isbn13(raw.isbn), legacy=clean(raw.legacy_barcode).replace(/[\s-]/g,'');
    const price=raw.list_price==null || String(raw.list_price).trim()==='' ? null : Number(raw.list_price);
    const cover=raw.cover ? originalCover(new URL(raw.cover,ref.url).href) : '';
    const rate=clean(raw.rating);
    const entry={...fields,work_title:clean(raw.work_title || parts.work_title).slice(0,240),volume_title:title.slice(0,240),
      edition_type:raw.edition_type || parts.edition_type,edition_label:raw.edition_label || parts.edition_label,print_variant:raw.print_variant || parts.print_variant,
      volume_number:clean(raw.volume_number || parts.volume_number),region:raw.region || ref.site.region,
      volume_sequence_format:raw.volume_sequence_format || parts.volume_sequence_format,
      published_volume_count:raw.published_volume_count ?? parts.published_volume_count,
      serial_status:raw.serial_status || parts.serial_status,
      authors:(Array.isArray(raw.authors)?raw.authors:[]).map(clean).filter(Boolean).slice(0,30),publisher:canonicalPublisher(raw.publisher).slice(0,160) || null,
      isbn:code || null,legacy_barcode:!code && /^\d{10,13}$/.test(legacy) && !/^(978|979)/.test(legacy) ? legacy : null,
      pub_date:raw.pub_date || null,description:raw.description ? String(raw.description).slice(0,20000) : null,
      format:parts.content_type==='manga' ? trimFormat(raw.spec || '') : null,
      audience_rating:/限制|18\s*禁/.test(rate)?'restricted':/[輔辅]導|[輔辅]导/.test(rate)?'guidance':/^普$|普通|普遍|全年齢/.test(rate)?'general':'pending',
      list_price:Number.isFinite(price) && price>=0 ? price : null,currency:price!=null ? raw.currency || null : null,
      notes:'资料来源：'+ref.url+(raw.spec ? '\n原站规格：'+clean(raw.spec).slice(0,500) : '')+(raw.metadata_note?'\n'+raw.metadata_note:'')};
    const warnings=[];
    if(!code) warnings.push(raw.isbn ? '书号校验失败，请核对' : entry.legacy_barcode ? '仅 EAN，可按 EAN 查书和入库' : 'ISBN 暂缺，可带封面本地入库');
    if(!cover) warnings.push('封面暂缺');
    if(!entry.description) warnings.push('简介暂缺');
    if(!entry.format) warnings.push('开本暂缺，请核对原站规格');
    return {id,title,entry,cover,coverAlternative:raw.cover_alternative || '',warnings,metadataRevision:1,ebook,blocked:raw.blocked || (ebook?'电子书不能作为实体卷册导出':parts.is_boxed?'盒装 / 多册合售不能当作一卷导出，请核对套装':''),fetchedAt:Date.now()};
  }
  function doubanIntro(doc) {
    let reading=false; const texts=[];
    for(const node of doc.querySelectorAll('h2,h3,.intro')) {
      if (/^H[23]$/.test(node.tagName)) reading=/内容简介|內容簡介/.test(node.textContent);
      else if(reading) {
        const copy=node.cloneNode(true);
        copy.querySelectorAll('script,style,a.j').forEach(n=>n.remove());
        copy.querySelectorAll('br').forEach(n=>n.replaceWith('\n'));
        copy.querySelectorAll('p,div').forEach(n=>{n.prepend('\n');n.append('\n');});
        texts.push(copy.textContent.split('\n').map(clean).filter(Boolean).join('\n'));
      }
    }
    return texts.sort((a,b)=>b.length-a.length)[0]?.slice(0,20000) || null;
  }
  function parseDouban(html,id,finalUrl=productUrl(id)) {
    const doc=new DOMParser().parseFromString(html,'text/html');
    const returned=new URL(finalUrl);
    if (['sec.douban.com','accounts.douban.com'].includes(returned.hostname) || /检测到有异常请求|安全验证/.test(doc.title) || doc.querySelector('form[action*="captcha"],input[name*="captcha"],form[action*="passport/login"]'))
      throw failure('豆瓣需要登录或人工验证。请打开该书处理后继续，不会自动重试验证页。','challenge');
    if(productId(finalUrl)!==id) throw failure('返回了其他豆瓣页面，未保存。');
    const canonical=doc.querySelector('link[rel="canonical"]')?.getAttribute('href') || doc.querySelector('meta[property="og:url"]')?.content;
    if(canonical && productId(canonical)!==id) throw failure('豆瓣书目编号与返回页面不一致，未保存。');
    const title=clean(doc.querySelector('h1 [property="v:itemreviewed"]')?.textContent || doc.querySelector('meta[property="og:title"]')?.content || doc.querySelector('h1')?.textContent).replace(/\s*\(豆瓣\)\s*$/,'');
    const info=doc.querySelector('#info');
    if(!title || !info) throw failure('豆瓣详情未加载完整，或当前是错误/验证页；未缓存，请打开图书核对。');
    const copy=info.cloneNode(true); copy.querySelectorAll('br').forEach(n=>n.replaceWith('|BCD_BR|'));
    const infoText=clean(copy.textContent).replaceAll('|BCD_BR|','\n');
    const labels='作者|译者|譯者|出版社|出版年|出版日期|页数|頁數|定价|定價|装帧|裝幀|ISBN|丛书|叢書|副标题|原作名|出品方|開本|开本|分級|分级';
    const fields={};
    const pattern=new RegExp('('+labels+')\\s*[:：]\\s*(.*?)(?=(?:'+labels+')\\s*[:：]|\\n|$)','g');
    for(const match of infoText.matchAll(pattern)) fields[match[1]]=clean(match[2]);
    const code=isbn13(fields.ISBN || doc.querySelector('meta[property="book:isbn"]')?.content);
    const region=regionFromIsbn(code);
    const priceText=fields['定价'] || fields['定價'] || '';
    const priceMatch=priceText.match(/\d[\d,]*(?:\.\d+)?/);
    const currency=/NT\s*\$|TWD|新?[台臺]幣/i.test(priceText) ? 'TWD' : /HK\s*\$|HKD|港[元幣币]/i.test(priceText) ? 'HKD' : /RMB|CNY|人民[币幣]/i.test(priceText) ? 'CNY' : /JPY|日[元圓圆]/i.test(priceText) ? 'JPY' : /US\s*\$|USD|美[元金]/i.test(priceText) ? 'USD' : {TW:'TWD',HK:'HKD',CN:'CNY',JP:'JPY'}[region] || null;
    const intro=doubanIntro(doc);
    const parts=enrichIdentity(identity(title),doc,intro), {ebook,...identityFields}=parts;
    const date=(fields['出版年'] || fields['出版日期'] || '').match(/^(\d{4})(?:[-/年](\d{1,2}))?(?:[-/月](\d{1,2}))?日?$/);
    const rating=fields['分級'] || fields['分级'] || '';
    const cover=[doc.querySelector('#mainpic a.nbg')?.getAttribute('href'),doc.querySelector('meta[property="og:image"]')?.content].map(originalCover).find(Boolean) || '';
    const entry={...identityFields,work_title:parts.work_title.slice(0,240),volume_title:title.slice(0,240),region,
      authors:(fields['作者'] || '').split(/\s*[/／、;；]\s*/).map(clean).filter(Boolean).slice(0,30),
      publisher:canonicalPublisher(fields['出版社']).slice(0,160) || null,isbn:code || null,legacy_barcode:null,
      pub_date:date ? date[1]+(date[2] ? '-'+date[2].padStart(2,'0') : '')+(date[3] ? '-'+date[3].padStart(2,'0') : '') : null,
      description:intro,list_price:priceMatch ? Number(priceMatch[0].replaceAll(',','')) : null,currency:priceMatch ? currency : null,
      format:trimFormat(fields['開本'] || fields['开本'] || ''),audience_rating:/限制/.test(rating) ? 'restricted' : /普/.test(rating) ? 'general' : /[輔辅]/.test(rating) ? 'guidance' : 'pending',
      notes:'资料来源：'+productUrl(id)};
    const warnings=[];
    if(!code) warnings.push('有效 ISBN 暂缺，不能按 ISBN 自动搜书');
    if(region==='UNKNOWN') warnings.push('地区待确认，请选择实际出版地区');
    if(!entry.description) warnings.push('简介暂缺');
    if(!entry.format) warnings.push('开本暂缺，不把装帧当开本');
    if(!cover) warnings.push('封面暂缺');
    return {id,title,entry,cover,warnings,blocked:ebook ? '电子书不能作为实体卷册导出' : parts.is_boxed ? '盒装 / 多册合售请在书库按套装核对，不能当作一卷导出' : '',fetchedAt:Date.now()};
  }
  function collectItems(doc,base=typeof location!=='undefined'?location.href:'https://www.books.com.tw/') {
    if(isBooksSearch(base)) {
      const found=new Map();
      for(const row of doc.querySelectorAll('.table-searchbox [id^="prod-itemlist-"],main [id^="prod-itemlist-"]')) {
        if(row.closest('#bcd-panel,#br_searchkw2item,.br_searchkw2item'))continue;
        const id=(row.id.match(/^prod-itemlist-(001\d+|E\d+)$/) || [])[1];
        const title=clean(row.querySelector('h4 a')?.textContent);
        if(id && title && !found.has(id))found.set(id,{id,title,selected:false,record:null,error:''});
      }
      return [...found.values()].slice(0,MAX_ITEMS);
    }
    const source=SITE_ADAPTERS.find(s=>s.hosts.includes(new URL(base).hostname));
    if(source) {
      const found=new Map();
      for(const row of source.collect?.(doc,base) || []) {
        const id=productId(new URL(row.url,base).href), title=clean(row.title);
        if(id && publisherRef(id)?.site.key===source.key && title && !found.has(id)) {
          const record=source.key==='sanmin' && row.html ? parsePublisher(row.html,id,productUrl(id),null,row.queryCode) : null;
          found.set(id,{id,title,selected:false,record,error:''});
        }
      }
      return [...found.values()].slice(0,MAX_ITEMS);
    }
    if (DOUBAN || doc.querySelector('.subject-list')) {
      const ids=new Map();
      for(const row of doc.querySelectorAll('#content .article .subject-list > li')) {
        const a=row.querySelector('h2 a[href],.info h2 a[href]');
        const id=productId(a?.getAttribute('href') || '');
        if(id.startsWith('douban-') && !ids.has(id)) ids.set(id,{id,title:clean(a.textContent),selected:false,record:null,error:''});
      }
      return [...ids.values()]; // No recommendations, automatic paging, or silent selection.
    }
    const checks = [...doc.querySelectorAll('input.prodCheck')];
    const ids = new Map(), titles = new Map();
    for (const a of doc.querySelectorAll('a[href*="/products/"]')) {
      if (a.closest('#bcd-panel')) continue;
      const id=productId(a.href), title=clean(a.querySelector('img')?.alt || a.textContent);
      if (id && title && !titles.has(id)) titles.set(id,title);
    }
    for (const check of checks) {
      const id = clean(check.value);
      if (!/^[A-Za-z0-9_-]+$/.test(id) || ids.has(id)) continue;
      const title = titles.get(id) || id;
      ids.set(id, {id, title, selected:check.checked, record:null, error:''});
    }
    return [...ids.values()]; // Never sweep recommendations or infer all when nothing is checked.
  }
  function publisherDirectoryId(site,url,legacy=false) {
    const u=new URL(url);u.hash='';u.pathname=u.pathname.replace(/\/page\/\d+\/?$/,'/').replace(/\/$/,'');
    for(const key of ['page','p','offset','limit','sort','order']) u.searchParams.delete(key);
    if(site.key==='kingstone') {
      const parts=u.pathname.split('/').filter(Boolean), filters=[];
      for(let i=3;i<parts.length;i+=2) filters.push([parts[i],parts[i+1] || '']);
      u.pathname='/'+parts.slice(0,3).concat(filters.sort((a,b)=>a[0].localeCompare(b[0])).flat()).join('/');
    }
    if(site.key==='eslite') for(const key of ['start','display','exp']) u.searchParams.delete(key);
    if(!legacy && site.key==='sanmin') u.searchParams.delete('pi');
    if(!legacy && site.key==='citicomics') for(const key of ['sort_by','order_by']) u.searchParams.delete(key);
    if(!legacy && site.key==='dala' && !u.searchParams.has('k')) u.searchParams.set('k','comic');
    if(site.key==='taaze') {
      for(const key of ['nowPage','layout']) u.searchParams.delete(key);
      for(const key of ['prodKind','prodCatId','catId','saleDiscStart','saleDiscEnd','salePriceStart','salePriceEnd','publishDateStart','publishDateEnd','prodRank','addMarkFlg','force','catFocus','orgFocus','mainCatFocus','catNmFocus','catIdFocus'])
        if(['','0'].includes(u.searchParams.get(key))) u.searchParams.delete(key);
    }
    u.searchParams.sort();return site.key+':'+u.pathname+u.search;
  }
  function booksSearchDirectoryId(value) {
    const u=new URL(value), parts=u.pathname.slice('/search/query/'.length).split('/').filter(Boolean), params=new URLSearchParams();
    // 博客来原站翻页会重排路径参数；页码、显示模式不改变搜索目录。
    for(let i=0;i<parts.length;i+=2) if(!['page','v'].includes(parts[i])) params.append(parts[i],parts[i+1] || '');
    for(const [key,val] of u.searchParams) if(!['page','v'].includes(key)) params.append(key,val);
    params.sort();return 'books-search:/search/query/?'+params.toString();
  }
  function legacyBooksSearchDirectoryId(value) {
    const u=new URL(value);u.hash='';u.pathname=u.pathname.replace(/\/page\/\d+(?=\/|$)/,'/page/1');u.searchParams.delete('page');u.searchParams.sort();return 'books-search:'+u.pathname+u.search;
  }
  function listingDirectoryId(value) {
    const u=new URL(value), site=SITE_ADAPTERS.find(s=>s.hosts.includes(u.hostname));
    return isBooksSearch(value) ? booksSearchDirectoryId(value) : site?.list(u) ? publisherDirectoryId(site,value) : '';
  }
  function nextListingPage(doc,base) {
    const key=listingDirectoryId(base);
    const controls=[...doc.querySelectorAll('a.jq-page-option[rel="next"]'),...doc.querySelectorAll('a[rel="next"],a.next[rel="change_page"][href],.PageNumBar .Next a[href],.ProductList-paginator a[href]:has(.fa-angle-right),.pagination-custom a[title="Next"][href],.pagination .pageNext a[href],.paginator .next a[href],.pagination .next a[href],a.next.page-numbers[href]')];
    for(const a of controls) {
      try {
        const u=new URL(a.getAttribute('href'),base);u.hash='';
        if(u.protocol==='https:' && !u.username && !u.password && !u.port && u.origin===new URL(base).origin && u.href!==base && listingDirectoryId(u.href)===key) return u.href;
      } catch { /* Only explicit links within the same listing are paging controls. */ }
    }
    return '';
  }
  async function loadNextListingPage() {
    if(state.busy || !state.nextListing) return;
    const target=state.nextListing;state.busy=true;state.stop=false;render();
    try {
      progress('正在加载下一页书目，已有勾选保留…');
      const res=await request(target);
      if(listingDirectoryId(res.url)!==state.directory.id) throw failure('分页返回了其他目录，请在网站核查。');
      const doc=new DOMParser().parseFromString(res.html,'text/html'), items=collectItems(doc,res.url);
      if(!items.length) throw failure('下一页未返回可识别书目，请在网站翻页后继续选择。');
      mergeItems(items,false);state.directory.done=[...new Set([...state.directory.done,res.url])];
      state.nextListing=nextListingPage(doc,res.url);
      if(state.directory.done.includes(state.nextListing)) state.nextListing='';
      ui.listDetails.open=true;
      progress(`已载入下一页 ${items.length} 项，新载入书目尚未勾选。请在下方选择所需卷册。`);
    } catch(err) {progress(err.message);}
    finally {state.busy=false;state.abort=null;render();}
  }
  const sortItems = items => [...items].sort((a,b) =>
    clean(a.entry.work_title).localeCompare(clean(b.entry.work_title), 'zh-Hant') ||
    (a.entry.sequence_role === 'supplement') - (b.entry.sequence_role === 'supplement') ||
    clean(a.entry.volume_number).localeCompare(clean(b.entry.volume_number), 'zh-Hant', {numeric:true}) || a.id.localeCompare(b.id));

  // STORE ZIP: no runtime CDN dependency. Both formats are read by the actual desktop decoders in tests.
  const encoder = new TextEncoder();
  const crcTable = Uint32Array.from({length:256}, (_,i) => { for (let j=0; j<8; j++) i = (i & 1) ? 0xedb88320 ^ (i >>> 1) : i >>> 1; return i >>> 0; });
  function crc32(bytes) { let crc = 0xffffffff; for (const b of bytes) crc = crcTable[(crc ^ b) & 255] ^ (crc >>> 8); return (crc ^ 0xffffffff) >>> 0; }
  async function zip(files, type = 'application/zip') {
    const chunks = [], central = []; let offset = 0, centralSize = 0;
    for (const [name, value] of files) {
      const bytes = typeof value === 'string' ? encoder.encode(value) : value instanceof Blob ? new Uint8Array(await value.arrayBuffer()) : value;
      const path = encoder.encode(name), crc = crc32(bytes);
      const head = new Uint8Array(30 + path.length), h = new DataView(head.buffer);
      h.setUint32(0,0x04034b50,true); h.setUint16(4,20,true); h.setUint16(6,0x800,true);
      h.setUint16(12,33,true); h.setUint32(14,crc,true); h.setUint32(18,bytes.length,true); h.setUint32(22,bytes.length,true); h.setUint16(26,path.length,true); head.set(path,30);
      const cd = new Uint8Array(46 + path.length), c = new DataView(cd.buffer);
      c.setUint32(0,0x02014b50,true); c.setUint16(4,20,true); c.setUint16(6,20,true); c.setUint16(8,0x800,true); c.setUint16(14,33,true);
      c.setUint32(16,crc,true); c.setUint32(20,bytes.length,true); c.setUint32(24,bytes.length,true); c.setUint16(28,path.length,true); c.setUint32(42,offset,true); cd.set(path,46);
      chunks.push(head,bytes); central.push(cd); offset += head.length + bytes.length; centralSize += cd.length;
      if (offset + centralSize + 22 > MAX_ZIP) throw failure('文件超过 128 MB，请取消部分勾选后分批导出。');
    }
    const tail = new Uint8Array(22), t = new DataView(tail.buffer);
    t.setUint32(0,0x06054b50,true); t.setUint16(8,files.length,true); t.setUint16(10,files.length,true); t.setUint32(12,centralSize,true); t.setUint32(16,offset,true);
    return new Blob([...chunks,...central,tail], {type});
  }
  // eslint-disable-next-line no-control-regex -- XML 1.0 forbids these control characters.
  const xml = v => String(v ?? '').replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,'').replace(/[&<>"']/g,c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'})[c]);
  const labelOf = (values,k) => values.find(([key]) => key === k)?.[1] || '';
  async function searchWorkbook(records) {
    const rows = [['作品名','地区','卷号','ISBN','出版社','漫画版本','卷册形式','作者','总卷数','连载状态','单卷标题','书籍类型','卷册角色','卷册类别','旧条码']];
    for (const {entry:e} of sortItems(records)) rows.push([e.work_title,labelOf(REGIONS,e.region),e.volume_number,e.isbn || '',e.publisher,
      labelOf(EDITIONS,e.edition_type),labelOf(FORMATS,e.volume_sequence_format),e.authors.join('、'),e.published_volume_count,
      labelOf(SERIAL,e.serial_status),e.volume_title,{manga:'漫画',novel:'小说',guidebook:'公式书',artbook:'画集'}[e.content_type],e.sequence_role === 'supplement' ? '附录' : '正篇',labelOf(KINDS,e.volume_kind),e.legacy_barcode || '']);
    const sheet = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" state="frozen"/></sheetView></sheetViews><cols>' +
      [32,10,12,21,18,14,14,24,12,14,42,14,14,18,21].map((w,i) => `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('') + '</cols><sheetData>' +
      rows.map((row,i) => `<row r="${i+1}" ht="28" customHeight="1">` + row.map((value,j) => `<c r="${String.fromCharCode(65+j)}${i+1}" t="inlineStr" s="${i ? 0 : 1}"><is><t xml:space="preserve">${xml(value)}</t></is></c>`).join('') + '</row>').join('') + `</sheetData><autoFilter ref="A1:O${rows.length}"/></worksheet>`;
    return zip([
      ['[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'],
      ['_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],
      ['xl/workbook.xml','<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="搜书清单" sheetId="1" r:id="rId1"/></sheets></workbook>'],
      ['xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'],
      ['xl/styles.xml','<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Microsoft JhengHei"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Microsoft JhengHei"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF34506E"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" xfId="0" numFmtId="49"/><xf fontId="1" fillId="2" borderId="0" xfId="0" numFmtId="49"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'],
      ['xl/worksheets/sheet1.xml',sheet]
    ], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }
  function recommendedScope(records, combine=false) {
    return !combine && records.length===1 && records[0].entry.volume_number && !records[0].entry.is_boxed ? 'volume' : 'work';
  }
  async function localPackage(records, covers, exportScope=recommendedScope(records)) {
    const files = [], groups = new Map(), items=[];
    const sha = async bytes => [...new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))].map(n=>n.toString(16).padStart(2,'0')).join('');
    if(exportScope==='volume' && records.length!==1) throw failure('单卷导出只能选择一册；多册请选择作品导出。');
    for(const record of sortItems(records)) {
      const id=await sha(encoder.encode('source:'+record.id)), blob=covers.get(record.id);
      const key=setKey(record.entry);
      if(!groups.has(key)) groups.set(key,{id:'group-'+(groups.size+1),work_title:record.entry.work_title});
      let cover=null;
      if(blob) {
        const ext={'image/jpeg':'jpg','image/png':'png','image/webp':'webp'}[blob.type];
        if(!ext) throw failure('封面格式不支持，请重新读取。');
        const picture=await createImageBitmap(blob);
        try {
          const file='covers/'+id+'.'+ext;
          cover={file,sha256:await sha(await blob.arrayBuffer()),bytes:blob.size,width:picture.width,height:picture.height};
          files.push([file,blob]);
        } finally {picture.close();}
      }
      const ref=publisherRef(record.id);
      const provenance={site:ref?.site.key || (record.id.startsWith('douban-') ? 'douban' : 'books'),
        product_id:record.id,source_title:record.title,work_title:identity(record.title).work_title,
        volume_basis:identity(record.title).volume_number ? 'title' : 'manual',url:productUrl(record.id)};
      items.push({id,group_id:groups.get(key).id,entry:record.entry,provenance,cover});
    }
    const manifest = JSON.stringify({format:'manga-library-local',version:2,export_scope:exportScope,work_groups:[...groups.values()],items});
    if (encoder.encode(manifest).length > 2*1024*1024) throw failure('资料超过 2 MB，请分批导出。');
    if (!records.length || records.length > MAX_ITEMS) throw failure('每包需要 1–300 本。');
    return zip([['manifest.json',manifest], ...files]);
  }
  async function jpeg(blob) {
    if (!blob.size || blob.size > 20*1024*1024) throw failure('图片为空或超过 20 MB。');
    const bytes=await blob.arrayBuffer(), head=new Uint8Array(bytes);
    const type=head[0]===255 && head[1]===216 && head[2]===255 ? 'image/jpeg' :
      head[0]===137 && head[1]===80 && head[2]===78 && head[3]===71 ? 'image/png' :
      String.fromCharCode(...head.slice(0,4))==='RIFF' && String.fromCharCode(...head.slice(8,12))==='WEBP' ? 'image/webp' : '';
    if(!type) throw failure('返回的不是可读取图片，或不是支持的 JPEG、PNG、WEBP；未写入建档包。');
    let image;
    try { image = await createImageBitmap(blob); } catch { throw failure('返回的不是可读取图片，未写入建档包。'); }
    try {
      if (image.width * image.height > 40_000_000 || Math.min(image.width,image.height) < 80) throw failure('图片像素异常，请打开商品检查。');
      return new Blob([bytes],{type});
    } finally { image.close(); }
  }
  function retryAfter(headers, now = Date.now()) {
    const raw = (String(headers || '').match(/^retry-after:\s*(.+)$/im) || [])[1];
    if (!raw) return 0;
    return /^\d+$/.test(raw.trim()) ? +raw*1000 : Math.max(0, Date.parse(raw)-now) || 0;
  }

  const volumeIdentity = item => {
    const e=item.record?.entry;
    return e ? JSON.stringify([workKey(e.work_title),e.region,publisherKey(e.publisher),e.edition_type,e.content_type,e.volume_kind || 'main',e.volume_number]) : '';
  };
  function preferPhysicalSelection(items) {
    const physical = new Set(items.filter(item=>item.record && !item.record.blocked && !item.record.ebook).map(volumeIdentity));
    return items.map(item=>item.record?.ebook && physical.has(volumeIdentity(item)) ? {...item,selected:false} : item);
  }

  // CommonJS exposes pure functions to the offline regression harness, not to the webpage.
  if (typeof module === 'object' && module.exports) {
    module.exports = {SITE_ADAPTERS,sanminCover,parsePublisher,parseBook,parseDouban,productId,productUrl,identity,numberText,isbn13,trimFormat,originalCover,collectItems,searchWorkbook,localPackage,jpeg,retryAfter,sortItems,inferSets,prepareSets,setKey,seriesPage,seriesUrl,canonicalPublisher,relatedSeries,supplementIds,directoryGaps,preferPhysicalSelection,booksSearchDirectoryId,citiVariants};
    return;
  }
  const state = {items:[], busy:false, stop:false, abort:null, nextAt:0, cooldownUntil:0, covers:new Map(), message:'核对选中的卷册，再读取资料。', warning:'', blockedUrl:'', directory:null, dirty:new Set(), auto:new Map(), currentIds:new Set(), selectionChanged:false, combine:false, combineApproved:''};
  const ui = {};
  const readStore = (key, fallback) => { try { return GM_getValue(key,fallback); } catch { return fallback; } };
  const cacheRead = id => {
    const hit=readStore(CACHE,{})[id];
    if(!hit || Date.now()-hit.fetchedAt>=TTL) return null;
    if(!id.startsWith('douban-') && !hit.metadataRevision && !hit.entry?.authors?.length) return null;
    const rec=structuredClone(hit), parts=identity(rec.title);
    rec.ebook=!DOUBAN && parts.ebook;
    if (rec.ebook && !id.startsWith('pub-')) {
      rec.entry.volume_title=stripEbookSuffix(rec.title);
      const sourceNote='（博客来无实体商品时采用电子书资料；缺失资料请人工补录）';
      if (!(rec.entry.notes || '').includes(sourceNote)) rec.entry.notes=(rec.entry.notes || '资料来源：'+productUrl(id))+sourceNote;
      const supportedId=/^001\d+$/.test(id) || /^E\d+$/i.test(id);
      rec.blocked=!supportedId ? '仅支持中文博客来商品' : parts.is_boxed ? '盒装 / 多册合售请在书库按套装核对，不能当作一卷导出' : '';
      rec.warnings=[...(rec.warnings || []).filter(w=>!w.includes('电子书')),'博客来当前仅提供电子书资料；缺失项目请人工补录'];
    }
    rec.entry.publisher=canonicalPublisher(rec.entry.publisher) || null;
    // Reparse cached titles with the new rules without downloading again.
    // Explicit per-volume draft edits are applied by the caller afterward.
    if(!id.startsWith('pub-')) Object.assign(rec.entry,{work_title:parts.work_title,volume_number:parts.volume_number || rec.entry.volume_number,
      edition_type:parts.edition_type,edition_label:parts.edition_label});
    return rec;
  };
  function cacheWrite(record) {
    const entries = Object.entries({...readStore(CACHE,{}), [record.id]:record}).filter(([,v]) => Date.now()-v.fetchedAt < TTL).sort((a,b) => b[1].fetchedAt-a[1].fetchedAt).slice(0,500);
    try { GM_setValue(CACHE,Object.fromEntries(entries)); } catch { state.warning='本次资料仍可导出，但缓存存储失败，关闭页面后可能需要重新读取。'; }
  }
  const selected = () => state.items.filter(item => item.selected);
  function preferPhysicalProducts() {
    if (state.selectionChanged) return;
    state.items=preferPhysicalSelection(state.items);
  }
  const bundle = item => identity(item.record?.title || item.title).is_boxed;
  function selectLoaded() {
    state.items.forEach(item=>{item.selected=ui.includeBundles.checked || !bundle(item);});
    state.selectionChanged=true;
  }
  function unreadSelectionHint() {
    const unselected=state.items.filter(item=>!item.selected && !bundle(item)).length;
    return DOUBAN && state.directory && unselected ? `另有 ${unselected} 本卷册未勾选，不会读取；要读整套，请在第 1 步“批量选择”中选择已载入整套。` : '';
  }
  function markEdit(item,field,value) {
    item.edits={...item.edits,[field]:value}; item.record.entry[field]=value; refreshAuto(); saveDraft();
  }
  function mergeItems(items,selectNew=true) {
    const existing=new Map(state.items.map(item=>[item.id,item]));
    for(const item of items) if(!existing.has(item.id)) {
      if(existing.size>=MAX_ITEMS) throw failure('目录超过 300 项，请缩小范围；尚未将目录标记为完整。');
      existing.set(item.id,{...item,selected:!selectNew || isSearchResults(location.href) ? false : state.selectionChanged ? item.selected : !bundle(item),record:cacheRead(item.id)});
    }
    state.items=[...existing.values()];
  }
  function saveDraft() {
    const draftId=state.directory?.id || productId(location.href);
    if(!draftId || !ui.work) return;
    const manual=Object.fromEntries([...state.dirty].map(key=>[key,ui[key]?.value || '']));
    const items=state.items.map(({id,title,selected,edits,review})=>({id,title,selected,edits,review}));
    try { GM_setValue('bcd-series-v1:'+draftId,{savedAt:Date.now(),directory:state.directory,items,manual,selectionChanged:state.selectionChanged,combine:state.combine,exportScope:ui.exportScope?.value || 'auto'}); }
    catch {state.warning='跨页草稿保存失败，当前页资料仍可导出，请暂勿刷新。';}
  }
  function refreshAuto() {
    const records=selected().filter(i=>i.record && !i.error).map(i=>i.record);
    const heading=clean(document.querySelector('h1')?.textContent);
    const forced=ui.grouping?.value==='same' ? clean(ui.work.value) || identity(heading).work_title || records[0]?.entry.work_title : '';
    state.prepared=prepareSets(records,heading,forced);
    if(ui.exportHint) ui.exportHint.textContent=recommendedScope(state.prepared,state.combine)==='volume'
      ? '建议导出单卷，可在 PC 作品的版本中导入；也可改为作品导出。'
      : '建议导出作品，将按作品和版本分别归组，请核对每册资料。';
    const directory=state.items.map(i=>({...i,record:i.record ? prepareSets([i.record],heading,forced)[0] : null}));
    state.auto=inferSets(state.prepared,directory,Boolean(state.directory?.complete));
    const one=state.auto.size===1 ? [...state.auto.values()][0] : null;
    const fields={work:'work_title',edition:'edition_type',sequence:'volume_sequence_format',total:'published_volume_count',serial:'serial_status',region:'region',authors:'authors',publisher:'publisher',content:'content_type'};
    for(const [key,field] of Object.entries(fields)) if(ui[key] && !state.dirty.has(key)) {
      const value=one?.defaults[field]; ui[key].value=Array.isArray(value) ? value.join('、') : value ?? '';
    }
    if(ui.inference) ui.inference.textContent=one ? `已识别为同一作品（${records.length} 册已读取）：`+(one.notes.join('；') || '已识别卷册形式与完结总数，请核对。') : records.length ? `识别到 ${state.auto.size} 组：不同地区、版本或附属册分别保留。若只是书名不同，可选择“按同一作品归组”并填写统一作品名。` : '读取详情后自动整理；缺少依据的资料仍可手填。';
    if(ui.gaps) ui.gaps.textContent=directoryGaps(state.items).join('\n') || '目录没有发现中间缺号；不代表网站已收齐全套。';
    renderWorkPreview();
  }
  function resolvedRecord(record) {
    const rec=structuredClone(record), e=rec.entry, inferred=state.auto.get(setKey(e))?.defaults || {};
    for(const field of ['work_title','publisher','authors']) {
      const value=inferred[field]; if(Array.isArray(value) ? value.length : value) e[field]=value;
    }
    if(e.sequence_role==='primary') for(const field of ['volume_sequence_format','published_volume_count','serial_status']) if(inferred[field]!=null) e[field]=inferred[field];
    const explicit=key=>state.dirty.has(key) && ui[key]?.value;
    if(explicit('work')) e.work_title=clean(ui.work.value);
    if(explicit('region')) e.region=ui.region.value;
    if(explicit('edition')) {e.edition_type=ui.edition.value;e.edition_label=labelOf(EDITIONS,ui.edition.value);}
    if(explicit('sequence') && e.sequence_role==='primary') e.volume_sequence_format=ui.sequence.value;
    if(explicit('serial') && e.sequence_role==='primary') e.serial_status=ui.serial.value;
    if(explicit('total') && e.sequence_role==='primary') e.published_volume_count=+ui.total.value;
    if(explicit('authors')) e.authors=ui.authors.value.split(/[、,，;；]/).map(clean).filter(Boolean);
    if(explicit('publisher')) e.publisher=canonicalPublisher(ui.publisher.value);
    if(explicit('content')) e.content_type=ui.content.value;
    return rec;
  }
  function renderWorkPreview() {
    if(!ui.workPreview) return;
    ui.combineReview.hidden=!state.combine;ui.workPreview.replaceChildren();
    const rows=state.prepared.map(resolvedRecord), groups=new Map();
    for(const row of sortItems(rows)) {const key=setKey(row.entry);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(row);}
    state.combineSignature=JSON.stringify([rows,selected().map(i=>[i.id,i.error || ''])]);
    ui.combineConfirm.checked=Boolean(state.combineApproved && state.combineApproved===state.combineSignature);
    ui.combineSummary.textContent=groups.size===1 ? `将整理为 1 个作品 · ${rows.length} 册。请核对下方卷号与商品身份。` : `当前识别 ${groups.size} 组。核对统一作品名，并取消其他作品或版本的卷册。`;
    for(const group of groups.values()) {
      const e=group[0].entry;
      el('h4',e.work_title,ui.workPreview);
      el('p',[labelOf(REGIONS,e.region),e.publisher || '出版社待确认',e.edition_label || '版本待确认'].join(' · '),ui.workPreview);
      const list=el('ol',null,ui.workPreview);
      for(const row of group) el('li',`${row.entry.volume_number || '卷号待确认'} · ${row.title}\n${row.entry.isbn || row.entry.legacy_barcode || '无书码'}`,list);
    }
  }
  async function addLinks(combine=false) {
    if(state.busy) return;
    let ids;
    const input=combine ? ui.combineLinks : ui.links;
    try {ids=supplementIds(input.value,DOUBAN);} catch(err) {progress(err.message);return;}
    if(new Set([...state.items.map(i=>i.id),...ids]).size>MAX_ITEMS) {progress('合计超过 300 项，请分批操作。');return;}
    if(combine) {state.combine=true;state.combineApproved='';}
    state.busy=true;state.stop=false;state.blockedUrl='';render();
    const baseline=selected().map(i=>i.record?.entry || {...identity(i.title),region:CURRENT_SITE?.region || (DOUBAN?'UNKNOWN':'TW')}).filter(e=>!e.is_boxed && e.sequence_role==='primary');
    let current=null, added=0, duplicate=0, review=0;
    try {
      for(const id of ids) {
        current=state.items.find(i=>i.id===id);
        if(current?.record) {if(combine && !current.record.blocked)current.selected=true;duplicate++;continue;}
        if(!current) {current={id,title:id,selected:false,record:null};state.items.push(current);}
        progress(`补入卷册 ${ids.indexOf(id)+1} / ${ids.length}：${id}`);
        current.error='';await fetchRecord(current);current.title=current.record.title;
        const e=current.record.entry;
        const code=e.isbn || e.legacy_barcode;
        const repeated=code && state.items.some(i=>i!==current && (i.record?.entry.isbn || i.record?.entry.legacy_barcode)===code);
        if(repeated && !combine) {
          state.items=state.items.filter(i=>i!==current);duplicate++;continue;
        }
        const same=baseline.length && baseline.some(b=>workKey(b.work_title)===workKey(e.work_title) && b.edition_type===e.edition_type && b.content_type===e.content_type && (!b.region || b.region==='UNKNOWN' || b.region===e.region) && (!b.publisher || !e.publisher || publisherKey(b.publisher)===publisherKey(e.publisher)));
        current.review=repeated ? '同 ISBN / EAN 的其他商品已在列表中；请核对后只保留一个。' : same && !current.record.blocked ? '' : combine ? '作品或版本需核对，请在第 2 步查看归组；书名不同可确认后统一作品名。' : '与当前作品 / 版本未能确认一致，暂未勾选；请在第 1 步核对后自行选择。';
        current.selected=Boolean(!repeated && !current.record.blocked && (same || combine));
        if(current.review) review++;
        added++;render();
      }
      preferPhysicalProducts();
      input.value='';progress(combine ? `已读取 ${added} 册；重复链接 / 已读商品 ${duplicate} 项。请核对统一作品名、版本和卷号。` : `已补入 ${added} 本；跳过 ${duplicate} 个重复商品 / ISBN。${review ? `${review} 本需回到第 1 步核对，未自动混入。` : '原勾选与手动修改保留。'}`);
      if(combine) {ui.showStep(2);await runAfterLinks();}
    } catch(err) {
      if(current && err.kind!=='pause') current.error=err.message;
      if(['challenge','login','cooldown'].includes(err.kind)) state.blockedUrl=productUrl(current.id);
      progress(err.message+' 已完成的卷册保留，可继续补入。');
    } finally {state.busy=false;state.abort=null;render();}
  }
  async function runAfterLinks() {
    // Read any previously selected listing rows through the same queue.
    for(const item of selected()) {
      if(state.stop) throw failure('已暂停分析，已读取资料保留。','pause');
      item.error='';await fetchRecord(item);
    }
  }
  async function loadDirectory() {
    if(state.busy || !state.directory) return;
    state.busy=true; state.stop=false; state.blockedUrl=''; render();
    let current='';
    try {
      const dir=state.directory;
      while(dir.pending.length) {
        if(state.stop) throw failure('已暂停目录读取，已加载的页与勾选保留。','pause');
        current=dir.pending[0]; progress(`读取丛书目录：已完成 ${dir.done.length} 页，已载入 ${state.items.length} 项…`);
        const res=await request(current), parsed=seriesPage(new DOMParser().parseFromString(res.html,'text/html'),res.url,dir.id);
        if(parsed.url!==current) throw failure('目录返回了其他页，未标记成功，请核对后继续。');
        mergeItems(parsed.items);
        dir.done=[...new Set([...dir.done,current])];
        dir.pending=[...new Set([...dir.pending.slice(1),...parsed.links])].filter(url=>!dir.done.includes(url));
        dir.complete=false; saveDraft(); render();
      }
      dir.complete=true; saveDraft(); progress(`目录已读完：${dir.done.length} 页 / ${state.items.length} 项（仅书目，不代表详情已读取）。请核对勾选范围，再读取所选详情。`);
    } catch(err) {
      state.directory.complete=false;
      state.blockedUrl=['challenge','login','cooldown'].includes(err.kind) ? current : '';
      progress(err.message); saveDraft();
    } finally {state.busy=false; state.abort=null; render();}
  }
  function progress(text) { state.message=text; ui.status.textContent=text + (state.warning ? '\n' + state.warning : ''); }
  function rememberCooldown(until) {
    state.cooldownUntil=Math.max(state.cooldownUntil,Number(readStore(COOLDOWN,0)) || 0,until);
    try { GM_setValue(COOLDOWN,state.cooldownUntil); }
    catch { state.warning='浏览器未能保存冷却时间；当前页仍会等待，请勿刷新或另开采集页绕过冷却。'; }
  }
  async function delay(ms) { const end=Date.now()+ms; while (Date.now()<end) { if (state.stop) throw failure('已暂停，已读取的资料保留。','pause'); await new Promise(r => setTimeout(r,Math.min(250,end-Date.now()))); } }
  async function request(url, blob=false) {
    const cooldown = Math.max(state.cooldownUntil,Number(readStore(COOLDOWN,0)) || 0);
    if (cooldown > Date.now()) throw failure('站方冷却中，可在 ' + new Date(cooldown).toLocaleTimeString() + ' 后点击继续。','cooldown');
    await delay(Math.max(0,state.nextAt-Date.now()));
    if (state.stop) throw failure('已暂停，已读取的资料保留。','pause');
    state.nextAt=Date.now()+GAP;
    return new Promise((resolve,reject) => {
      let finished=false, timer;
      const finish=callback=>value=>{
        if(finished)return;
        finished=true;clearTimeout(timer);state.abort=null;callback(value);
      };
      const handlers={
        onload:finish(res => {
          state.abort=null;
          const body=res.responseText || '';
          if ([429,484,503].includes(res.status)) {
            const ms=Math.max(retryAfter(res.responseHeaders),60000);
            rememberCooldown(Date.now()+ms);
            reject(failure('HTTP '+res.status+'：已暂停，最早 '+new Date(Date.now()+ms).toLocaleTimeString()+' 后继续。','cooldown',ms));
          } else if ([401,403,418].includes(res.status)) reject(failure('站方拒绝当前访问或要求登录/验证（HTTP '+res.status+'）。请打开该书核查后继续。','challenge'));
          else if (res.status>=200 && res.status<300) {
            if(blob && !originalCover(res.finalUrl || url)) reject(failure('封面跳转到其他来源，未保存。','challenge'));
            else if(blob && !/^image\/(?:jpeg|png|webp|gif)(?:;|$)/i.test(res.response?.type || '')) reject(failure('封面未返回可识别的图片，未保存。','invalid_image'));
            else resolve(blob ? res.response : {html:body,url:res.finalUrl || url});
          }
          else reject(failure('HTTP '+res.status, blob && [404,410].includes(res.status) ? 'missing_image' : res.status>=500 ? 'transient' : 'error'));
        }),onerror:finish(() => reject(failure('网络请求失败，请检查网络与油猴站点权限。','transient'))),
        ontimeout:finish(() => reject(failure('请求超时。','transient'))),
        onabort:finish(() => reject(failure('已暂停，已读取的资料保留。','pause')))};
      // Same-origin Books HTML uses the current tab's browser session. Images
      // and cross-origin requests retain the script manager's existing transport.
      if(!blob && location.hostname==='www.books.com.tw' && new URL(url).origin===location.origin) {
        const controller=new AbortController();
        state.abort=()=>{controller.abort();handlers.onabort();};
        timer=setTimeout(()=>{controller.abort();handlers.ontimeout();},30000);
        fetch(url,{credentials:'include',signal:controller.signal}).then(async res=>{
          handlers.onload({status:res.status,responseText:await res.text(),finalUrl:res.url,
            responseHeaders:'Retry-After: '+(res.headers.get('retry-after') || '')});
        }).catch(handlers.onerror);
      } else {
        const xhr=GM_xmlhttpRequest({method:'GET',url,timeout:30000,headers:{Referer:CURRENT_SITE ? location.origin+'/' : DOUBAN ? 'https://book.douban.com/' : 'https://www.books.com.tw/'},...(blob ? {responseType:'blob'} : {}),...handlers});
        if(!finished)state.abort=()=>xhr.abort();
      }
    });
  }
  async function fetchCoverBlob(record) {
    const urls=[...new Set([record.cover,record.coverAlternative].filter(url=>originalCover(url)))];
    for(const url of urls) {
      try {
        const blob=await request(url,true);
        let image;try {image=await createImageBitmap(blob);} catch {throw failure('封面图片无法解码。','invalid_image');}
        const valid=image.width>=80 && image.height>=80 && image.width*image.height<=40_000_000;image.close();
        if(!valid)throw failure('封面图片尺寸无效。','invalid_image');
        record.cover=url;return blob;
      } catch(error) {
        if(!['missing_image','invalid_image'].includes(error.kind))throw error;
      }
    }
    return null;
  }
  async function fetchRecord(item) {
    if (item.record) return;
    if(location.hostname==='www.books.com.tw' && productId(location.href)===item.id) {
      // A verification page may have finished loading since the toolbox booted.
      item.record=readCurrentProduct(item.id);
      cacheWrite(item.record);Object.assign(item.record.entry,item.edits || {});return;
    }
    const hit=cacheRead(item.id);
    if (hit) { item.record=structuredClone(hit); Object.assign(item.record.entry,item.edits || {}); return; }
    for (let attempt=0; attempt<2; attempt++) {
      try {
        const ref=publisherRef(item.id);
        if(ref?.site.detailApi) {
          const apiUrl=ref.site.detailApi(new URL(ref.url)), res=await request(apiUrl);
          if(res.url!==apiUrl) throw failure('官网书目接口跳转到其他地址，未保存。请在本站核查。');
          let data;try {data=JSON.parse(res.html);} catch {throw failure('官网书目接口未返回有效资料，请在本站核查。');}
          item.record=parsePublisher('',item.id,ref.url,data);
        } else {
          const res=await request(productUrl(item.id)); item.record=parseBook(res.html,item.id,res.url);
        }
        cacheWrite(item.record); Object.assign(item.record.entry,item.edits || {}); return;
      } catch (err) {
        if (err.kind !== 'transient' || attempt) throw err;
        progress('网络暂时异常，3 秒后重试当前一本…'); await delay(3000);
      }
    }
  }
  function el(tag, text, parent) { const n=document.createElement(tag); if (text != null) n.textContent=text; parent?.appendChild(n); return n; }
  function button(text, action, parent) { const n=el('button',text,parent); n.type='button'; n.addEventListener('click',action); return n; }
  function selectControl(label, values, value, parent, change) {
    const box=el('label',label,parent), select=el('select',null,box);
    select.setAttribute('aria-label',label);
    for (const [key,text] of values) { const opt=el('option',text,select); opt.value=key; }
    select.value=value; if (change) select.addEventListener('change',() => change(select.value)); return select;
  }
  function inputControl(label, value, parent, change) {
    const box=el('label',label,parent), input=el('input',null,box); input.type='text'; input.value=value || '';
    if (change) input.addEventListener('input',() => change(input.value)); return input;
  }
  function link(text,url,parent) { const a=el('a',text,parent); a.href=url; a.target='_blank'; a.rel='noopener noreferrer'; return a; }
  function searchSelectionRows() {
    if(!isSearchResults(location.href)) {
      if(!CURRENT_SITE || productId(location.href) || !CURRENT_SITE.list(new URL(location.href))) return [];
      const ids=new Set(collectItems(document).map(item=>item.id)), seen=new Set();
      const anchors=CURRENT_SITE.key==='citicomics' ? document.querySelectorAll('.ProductList-list .product-item a.Product-item[href]') : CURRENT_SITE.key==='dala' ? document.querySelectorAll('main .detail-group a[href*="/Publication/Detail"]') : CURRENT_SITE.key==='dart' ? document.querySelectorAll('.grid-link__container .product_grid-item a[href*="/products/"]') : CURRENT_SITE.key==='kodansha' ? document.querySelectorAll('[data-scroll-anchor="publications"] a[href*="/comic/products/"]') : CURRENT_SITE.key==='shogakukan' ? document.querySelectorAll('#main .bookBox01 p.ttl a[href]') : document.querySelectorAll('a[href]');
      return [...anchors].flatMap(a=>{
        if(a.closest('#bcd-panel')) return [];
        const id=productId(a.href);
        if(!ids.has(id) || seen.has(id)) return [];
        seen.add(id);
        const row=(CURRENT_SITE.key==='kodansha' ? a.closest('li') : null) || a.closest('li.product,.productBox,.single-item,.product-thumb,.product_grid-item,.ProductList-item,.product-item,li.clearfix') || a.parentElement;
        return [{row,container:CURRENT_SITE.key==='chingwin' ? row.querySelector('.box-front .nameBox') || row : row,id}];
      });
    }
    const rule=SEARCH_CHOICES[CURRENT_SITE?.key];
    if(rule) return [...document.querySelectorAll(rule.rows)].map(row=>{
      const href=(row.matches('a[href]') ? row : row.querySelector(rule.link))?.getAttribute('href');
      return {row,container:row.querySelector(rule.container || ':scope') || row,id:href ? productId(new URL(href,location.href).href) : ''};
    });
    return [...document.querySelectorAll('.table-searchbox [id^="prod-itemlist-"],main [id^="prod-itemlist-"]')]
      .filter(row=>!row.closest('#bcd-panel,#br_searchkw2item,.br_searchkw2item'))
      .map(row=>({row,container:row,id:(row.id.match(/^prod-itemlist-(001\d+|E\d+)$/) || [])[1]}));
  }
  function renderSearchSelection() {
    const items=new Map(state.items.map(item=>[item.id,item]));
    for(const entry of searchSelectionRows()) {
      const {row,id}=entry;let container=entry.container;
      const item=items.get(id);
      if(!item) continue;
      if(row.matches('a[href]')) {
        let shell=row.parentElement.matches('.bcd-card-shell') ? row.parentElement : null;
        if(!shell) {
          shell=el('div');shell.className=row.className.replace(/\bproduct-item\b/,'bcd-card-shell');
          for(const attr of row.attributes) if(attr.name.startsWith('data-v-')) shell.setAttribute(attr.name,attr.value);
          row.before(shell);shell.appendChild(row);
        }
        container=shell;
      }
      let label=container.querySelector('.bcd-search-choice');
      if(!label) {
        label=el('label'); label.className='bcd-search-choice'; container.prepend(label);
        const check=el('input',null,label); check.type='checkbox'; check.setAttribute('aria-label','选入书库：'+item.title);
        el('span','选入书库',label);
        check.addEventListener('change',()=>{
          item.selected=check.checked; state.selectionChanged=true;
          progress(`已选 ${selected().length} 本，请读取或核对资料后导出。`); render();
        });
      }
      const check=label.querySelector('input'); check.checked=item.selected; check.disabled=state.busy;
      label.querySelector('span').textContent=item.selected?'已选入书库':'选入书库';
      label.dataset.selected=String(item.selected);
    }
  }
  function render() {
    refreshAuto();
    renderSearchSelection();
    ui.list.replaceChildren(); const scope=selected();
    const ready=scope.filter(item=>item.record && !item.error).length;
    ui.badge.textContent=`已选 ${scope.length}`;
    ui.listSummary.textContent=`查看 / 调整卷册选择与逐本资料 · ${scope.length} 已选`;
    ui.count.textContent=`已选 ${scope.length} / 已载入 ${state.items.length} 项`+(state.directory ? ` · ${state.directory.complete?'目录已读完（仅书目）':'目录未读完'}` : '')+`\n已选详情已读 ${ready} / ${scope.length} 本 · 待读 ${scope.length-ready} 本`;
    if(unreadSelectionHint()) ui.count.textContent+='\n'+unreadSelectionHint();
    for (const item of state.items.filter(item=>(!ui.onlySelected?.checked || item.selected) && (!ui.variantBook?.value || item.title.startsWith(ui.variantBook.value+' ')))) {
      const row=el('div',null,ui.list); row.className='book';
      const label=el('label',null,row), check=el('input',null,label); check.type='checkbox'; check.checked=item.selected; check.disabled=state.busy;
      const rec=item.record; el('span',rec?.title || item.title,label);
      check.addEventListener('change',() => { item.selected=check.checked; state.selectionChanged=true; progress(`已选 ${selected().length} 本，请读取或核对资料后导出。`); render(); });
      link('打开商品',productUrl(item.id),row);
      if(bundle(item)) el('p','合集 / 盒装 / 合售：不计正篇卷数；可勾选核对，但不能当单卷导出。',row).className='warning';
      if (item.error) el('p',item.error,row).className='warning';
      if (item.review) el('p',item.review,row).className='warning';
      if (!rec) { el('p','尚未读取详情',row); continue; }
      if (rec.blocked) { el('p',rec.blocked,row).className='warning'; continue; }
      if (rec.ebook) el('p','电子书资料作为实体书资料替代；缺失项目请人工补录。',row).className='warning';
      el('p',`卷号：${rec.entry.volume_number || '待确认'} · ${rec.entry.edition_label || '版本待确认'}`,row);
      const details=el('details',null,row); el('summary','核对 / 修改本卷资料 · '+(rec.entry.isbn || rec.entry.legacy_barcode || '无书号'),details);
      const edit=el('div',null,details); edit.className='fields';
      const vol=inputControl('卷号',rec.entry.volume_number,edit,value => {markEdit(item,'volume_number',value);}); vol.maxLength=40; vol.placeholder='请核对';
      if(!rec.entry.volume_number && rec.entry.sequence_role==='primary') button('确认这是全一册',()=>{
        for(const [key,value] of Object.entries({volume_number:'全',volume_sequence_format:'single',published_volume_count:1,serial_status:'completed'})) markEdit(item,key,value);
        render();
      },row).dataset.idle='true';
      selectControl('卷册类别',KINDS,rec.entry.volume_kind,edit,value => {
        markEdit(item,'volume_kind',value); markEdit(item,'sequence_role',value==='main'?'primary':'supplement');
        if(['guidebook','novel','artbook'].includes(value)) markEdit(item,'content_type',value);
        render();
      });
      selectControl('内容类型',CONTENTS,rec.entry.content_type,edit,value=>{markEdit(item,'content_type',value);render();});
      selectControl('卷册角色',[['primary','正篇'],['supplement','附属册']],rec.entry.sequence_role,edit,value=>{markEdit(item,'sequence_role',value);render();});
      inputControl('作品名',rec.entry.work_title,details,v => {markEdit(item,'work_title',v);}).maxLength=240;
      if(DOUBAN) selectControl('出版地区',REGIONS,rec.entry.region,details,v=>{markEdit(item,'region',v);render();});
      const fields={作者:rec.entry.authors.join('、'),出版社:rec.entry.publisher,版本:rec.entry.edition_label,出版日期:rec.entry.pub_date,
        ...(DOUBAN ? {地区:labelOf(REGIONS,rec.entry.region)} : {}),
        分级:{general:'普通级',restricted:'限制级',guidance:'辅导级',pending:'待核实'}[rec.entry.audience_rating],开本:rec.entry.format,
        定价:rec.entry.list_price == null ? '' : rec.entry.list_price+' '+(rec.entry.currency || '币种待确认'),简介:rec.entry.description};
      for (const [key,value] of Object.entries(fields)) el('p',key+'：'+(value || '暂缺'),details);
      if (rec.warnings.length) el('p',rec.warnings.join('；'),details);
    }
    ui.panel.querySelectorAll('input,select,textarea,button[data-idle]').forEach(n => {n.disabled=state.busy;});
    if(ui.nextListing) ui.nextListing.disabled=state.busy || (!state.nextListing && CURRENT_SITE?.key!=='eslite');
    ui.stop.hidden=!state.busy;
    if(ui.back) {ui.back.hidden=state.busy || state.step===1;ui.next.hidden=state.busy || state.step===3;}
    ui.blocked.hidden=!state.blockedUrl; ui.blocked.href=state.blockedUrl || '#';
    progress(state.message); saveDraft();
  }
  function checkedRecords() {
    refreshAuto();
    const items=selected();
    if (!items.length || items.length>MAX_ITEMS) throw failure('请勾选 1–300 本；没有勾选时不会默认导出全部。');
    const total=ui.total.value.trim();
    if (total && (!/^\d+$/.test(total) || +total<1 || +total>300)) throw failure('总卷数应为 1–300 的整数，不确定时留空。');
    const records=items.map((item,index) => {
      if (!item.record || item.error) throw failure('还有未读成功的书，请先继续读取，或取消勾选失败项。');
      if (item.record.blocked) throw failure(item.record.title+'：'+item.record.blocked+'。请取消勾选。');
      const rec=resolvedRecord(state.prepared[index]), e=rec.entry;
      const explicit=key=>state.dirty.has(key) && ui[key]?.value;
      const fixed={single:1,upper_lower:2,upper_middle_lower:3}[e.volume_sequence_format];
      if (fixed && e.sequence_role==='primary') {
        if (explicit('total') && +total!==fixed) throw failure('卷册形式与总卷数不一致。全一卷 / 上下 / 上中下分别为 1 / 2 / 3 卷。');
        e.published_volume_count=fixed;
      }
      e.volume_number=clean(e.volume_number);
      if (!clean(e.work_title) || !e.volume_number) throw failure(rec.title+'：请补齐作品名和卷号。');
      if (e.sequence_role==='primary' && fixed) {
        const allowed={single:['全','1','全一卷','一本完'],upper_lower:['上','下','1','2'],upper_middle_lower:['上','中','下','1','2','3']}[e.volume_sequence_format];
        if (!allowed.includes(e.volume_number)) throw failure(rec.title+'：卷号与卷册形式不一致，请核对。');
        if (e.volume_sequence_format==='single') e.volume_number='全';
        else if (!/^\d+$/.test(e.volume_number)) e.volume_number=String(allowed.indexOf(e.volume_number)+1);
      }
      if (e.sequence_role==='primary' && ['上','中','下'].includes(e.volume_number) && e.volume_sequence_format==='numeric') throw failure(rec.title+'：请选择“上下”或“上中下”，不要猜总卷数。');
      return rec;
    });
    if(state.combine) {
      if(records.length<2) throw failure('多册合成作品需要至少 2 册，请回到第 1 步选择卷册或补入链接。');
      if(new Set(records.map(r=>setKey(r.entry))).size!==1) throw failure('仍有不同作品、出版社或版本。请核对统一作品名，并回到第 1 步取消其他组的卷册。');
      if(!ui.combineConfirm.checked || state.combineApproved!==state.combineSignature) throw failure('请先在第 2 步查看作品与卷册预览，勾选确认；修改资料后需要重新核对。');
    }
    const codes=new Set(), identities=new Set();
    for (const {entry:e} of records) {
      const key=JSON.stringify([e.work_title,e.region,e.publisher,e.edition_type,e.volume_kind,e.volume_number]);
      const code=e.isbn || e.legacy_barcode;
      if (identities.has(key) || code && codes.has(code)) throw failure('勾选项有相同卷号 / ISBN / EAN 的不同商品，请先取消重复版本。');
      identities.add(key); if(code) codes.add(code);
    }
    return records;
  }
  function save(blob,name) {
    const href=URL.createObjectURL(blob), a=el('a',null,document.body); a.href=href; a.download=name; a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(href),60000);
  }
  // eslint-disable-next-line no-control-regex -- Windows filenames cannot contain control characters.
  const filename = text => text.replace(/[<>:"/\\|?*\x00-\x1f]/g,'-').slice(0,90) || '博客来';
  async function run(mode) {
    if (state.busy) return;
    const scope=selected();
    if (!scope.length || scope.length>MAX_ITEMS) {progress('请勾选 1–300 本；“全选本页”只选择当前页面列出的书。'); return;}
    state.busy=true; state.stop=false; state.blockedUrl=''; render();
    let current=null;
    try {
      if (mode==='read') {
        let done=0, lastRender=0;
        for (const item of scope) {
          current=item; if(state.stop) throw failure('已暂停，已读取的资料保留。','pause');
          item.error=''; progress(`读取 ${done+1} / ${scope.length}：${item.title}`);
          await fetchRecord(item); done++;
          // Cached rows can finish synchronously: avoid rebuilding hundreds of
          // editors per row, and give pause/paint an event-loop turn.
          if(Date.now()-lastRender>150 || done===scope.length) {
            render(); lastRender=Date.now(); await new Promise(resolve=>setTimeout(resolve,0));
          }
        }
        preferPhysicalProducts();
        progress(`已读取 ${done} 本（本次勾选范围）。${unreadSelectionHint()}请核对作品名、卷号、版本和总卷数，再选择导出方式。`);
      } else {
        const records=checkedRecords();
        if (mode==='search') {
          const missing=records.filter(r => !r.entry.isbn).length;
          if (missing && !confirm(`${missing} 本没有有效 ISBN，清单会保留这些行，但书库不能自动联网核查。仍要导出吗？`)) { progress('已取消导出，资料保留。'); return; }
          const workbook=await searchWorkbook(records);
          if(state.stop) throw failure('已暂停，未导出文件。','pause');
          save(workbook,filename(ui.work.value || records[0].entry.work_title)+'-搜书清单.xlsx');
          progress(`已导出 ${records.length} 本搜书清单（不含图片）。到书库“搜书 → 导入 ISBN 搜书清单”选择此文件。`);
        } else {
          let done=0, size=0; const covers=new Map(), missing=[];
          for (const rec of records) {
            current=scope.find(item => item.id===rec.id);
            if(state.stop) throw failure('已暂停；本页已下载的封面保留，点击导出可继续。','pause');
            progress(`下载封面 ${++done} / ${records.length}：${rec.title}`);
            if (!rec.cover) {missing.push(rec.title); continue;}
            let blob=state.covers.get(rec.id);
            if (!blob) {
              const original=await fetchCoverBlob(rec);
              if(!original){missing.push(rec.title);continue;}
              blob=await jpeg(original); state.covers.set(rec.id,blob);
              let held=[...state.covers.values()].reduce((sum,b)=>sum+b.size,0);
              for(const [key,value] of state.covers) {if(held<=MAX_ZIP) break; state.covers.delete(key); held-=value.size;}
            }
            covers.set(rec.id,blob); size+=blob.size;
            if (size>MAX_ZIP-2*1024*1024) throw failure('图片合计过大，请分批勾选导出。');
          }
          if (missing.length && !confirm(`${missing.length} 本在商品页没有可用封面，将以“暂缺”入库：\n${missing.slice(0,8).join('\n')}\n仍要导出吗？`)) { progress('已取消导出，资料与已下载封面保留。'); return; }
          const purpose=ui.exportScope.value==='auto' ? recommendedScope(records,state.combine) : ui.exportScope.value;
          const packet=await localPackage(records,covers,purpose);
          if(state.stop) throw failure('已暂停，未导出文件；已下载封面保留。','pause');
          save(packet,filename(ui.work.value || records[0].entry.work_title)+(purpose==='volume' ? '-单卷建档.zip' : '-作品建档.zip'));
          progress(`已导出 ${records.length} 本资料、${covers.size} 张本地封面。到“我的书库 → 导入本地建档包”预览后确认。`);
        }
      }
    } catch (err) {
      if (current && mode==='read' && err.kind!=='pause') current.error=err.message;
      if (['login','challenge','cooldown'].includes(err.kind)) {
        state.blockedUrl=current ? productUrl(current.id) : '';
        if (err.delay) rememberCooldown(Date.now()+err.delay);
      }
      progress(err.message);
    } finally {state.busy=false; state.abort=null; render();}
  }
  async function downloadOriginal(record) {
    if(state.busy) return;
    state.busy=true; state.stop=false; render();
    try {
      const blob=await fetchCoverBlob(record);
      if(!blob)throw failure('金石堂和三民目前都没有可用封面。');
      const image=await createImageBitmap(blob); image.close(); // Never save a login/error page as a cover.
      const ext={'image/jpeg':'jpg','image/png':'png','image/webp':'webp','image/gif':'gif'}[blob.type.split(';')[0]] || record.cover.split('.').pop();
      save(blob,filename(record.title)+'.'+ext); progress('已保存封面原始文件；未重绘，也未重新压缩。');
    } catch(err) {progress('封面下载失败：'+err.message);}
    finally {state.busy=false; render();}
  }
  function readCurrentProduct(id) {
    const snapshot=document.documentElement.cloneNode(true);
    snapshot.querySelector('#bcd-panel')?.remove();
    // Some script managers inject inline source. Challenge-pattern literals
    // in this userscript are code, not site evidence.
    snapshot.querySelectorAll('script').forEach(node=>{if(node.textContent.includes('// @namespace    local.books.cover'))node.remove();});
    return parseBook(snapshot.outerHTML,id,location.href);
  }
  function boot() {
    if(document.getElementById('bcd-panel')) return;
    const variants=CURRENT_SITE?.key==='citicomics' && productId(location.href) ? citiVariants(document,location.href) : [];
    const id=variants.length ? '' : productId(location.href), booksSearch=isBooksSearch(location.href), searchResults=isSearchResults(location.href), isSeries=variants.length>0 || (CURRENT_SITE ? CURRENT_SITE.list(new URL(location.href)) : DOUBAN ? /^\/series\/\d+\/?$/.test(location.pathname) : booksSearch || location.hostname==='activity.books.com.tw' && location.pathname.startsWith('/crosscat/show/'));
    if(variants.length>MAX_ITEMS) throw failure('本页书目／期数组合超过 300 项，请缩小范围后采集。');
    if(!id && !isSeries) return;
    state.nextListing=isSeries ? nextListingPage(document,location.href) : '';
    if(id) {
      const item={id,title:id,selected:true,record:null,error:''};
      try {item.record=readCurrentProduct(id); item.title=item.record.title; cacheWrite(item.record); state.message='已读取当前商品页，无需重复联网。请核对卷号与资料后导出。';}
      catch(err) {item.error=err.message; state.message=err.message; state.blockedUrl=productUrl(id);}
      state.items=[item];
      if(CURRENT_SITE?.detailApi && !item.record) {item.error='';state.blockedUrl='';state.message='本页使用官网公开书目接口，请点击“读取 / 继续读取所选资料”，不会提前联网。';}
    } else {
      const current=variants.length ? variants.map(row=>{const id=productId(row.url), record=readCurrentProduct(id);cacheWrite(record);return {id,title:row.title,selected:new URL(location.href).searchParams.get('bcd_variant')===row.variant.key,record,error:''};}) : collectItems(document); state.currentIds=new Set(current.map(i=>i.id));
      state.items=current.map(item=>({...item,selected:variants.length ? item.selected : !searchResults && !bundle(item),record:cacheRead(item.id)}));
      if(variants.length)state.message=`已识别本页 ${variants.length} 个书目／期数组合，请逐项勾选；不同作品与版本分别归组，不代表全套完整。`;
      if(DOUBAN) {
        const sid=location.pathname.match(/\d+/)[0];
        try {
          const parsed=seriesPage(document,location.href,sid);
          state.directory={id:sid,done:[parsed.url],pending:parsed.links.filter(u=>u!==parsed.url),complete:parsed.links.every(u=>u===parsed.url)};
        } catch(err) {state.message=err.message;}
      } else state.directory={id:variants.length ? 'citicomics-variants:'+productId(citiBaseUrl(location.href)) : booksSearch ? booksSearchDirectoryId(location.href) : CURRENT_SITE ? publisherDirectoryId(CURRENT_SITE,location.href) : 'books-'+location.pathname.split('/').filter(Boolean).pop(),done:[location.href],pending:[],complete:!booksSearch && !CURRENT_SITE && !document.querySelector('.paginator a[href],a[rel="next"]') && !/部分(?:版本或)?(?:集數|集数|冊|册).*無法|部分.*无法/.test([...document.querySelectorAll('h1,h2,h3')].map(n=>clean(n.textContent)).join(' '))};
      if(state.directory) {
        // 升级时迁移当前地址的旧草稿；已有新草稿优先，避免恢复已取消的勾选。
        const legacyId=variants.length ? '' : booksSearch ? legacyBooksSearchDirectoryId(location.href) : CURRENT_SITE ? publisherDirectoryId(CURRENT_SITE,location.href,true) : '';
        const saved=readStore('bcd-series-v1:'+state.directory.id,null) ?? (legacyId && legacyId!==state.directory.id ? readStore('bcd-series-v1:'+legacyId,null) : null);
        if(saved && Date.now()-saved.savedAt<86400000 && Array.isArray(saved.items) && saved.items.length<=MAX_ITEMS) {
          const old=saved.items.filter(i=>typeof i.id==='string' && productId(productUrl(i.id))===i.id && sameSiteId(i.id)).map(i=>({...i,record:cacheRead(i.id),error:''}));
          state.items=old; mergeItems(current,DOUBAN);
          if(DOUBAN && saved.directory?.id===state.directory.id) {
            const order=new URL(state.directory.done[0]).searchParams.get('order');
            const valid=values=>(Array.isArray(values)?values:[]).map(u=>seriesUrl(u,location.href,state.directory.id)).filter(u=>u && new URL(u).searchParams.get('order')===order);
            const done=[...new Set([...valid(saved.directory.done),...state.directory.done])];
            const pending=[...new Set([...valid(saved.directory.pending),...state.directory.pending])].filter(u=>!done.includes(u));
            state.directory={id:state.directory.id,done,pending,complete:!pending.length};
          }
          state.restoredManual=saved.manual || {};
          state.restoredExportScope=saved.exportScope;
          state.combine=saved.combine===true;
          state.selectionChanged=saved.selectionChanged == null ? true : saved.selectionChanged === true;
          for(const item of state.items) if(item.record) Object.assign(item.record.entry,item.edits || {});
          state.message='已恢复本丛书的跨页勾选与手动修改，可继续加载目录或读取详情。';
        }
      }
    }
    if(id) {
      const saved=readStore('bcd-series-v1:'+id,null);
      if(saved && Date.now()-saved.savedAt<86400000 && Array.isArray(saved.items) && saved.items.length<=MAX_ITEMS) {
        const current=state.items[0];
        state.items=saved.items.filter(i=>typeof i.id==='string' && productId(productUrl(i.id))===i.id && sameSiteId(i.id)).map(i=>({...i,record:i.id===id?current.record:cacheRead(i.id)}));
        if(!state.items.some(i=>i.id===id)) state.items.unshift(current);
        for(const item of state.items) if(item.record) Object.assign(item.record.entry,item.edits || {});
        state.restoredManual=saved.manual || {};state.selectionChanged=saved.selectionChanged===true;
        state.restoredExportScope=saved.exportScope;
        state.combine=saved.combine===true;
      }
    }
    preferPhysicalProducts();
    const panel=el('section',null,document.body); panel.id='bcd-panel'; ui.panel=panel; panel.setAttribute('aria-label',SITE_NAME+'书库导出工具箱');
    const style=el('style',null,panel); style.textContent=`
      #bcd-panel{position:fixed;right:16px;bottom:16px;width:min(420px,calc(100vw - 24px));max-height:90vh;z-index:2147483000;background:#1f2430;color:#e8ecf3;border-radius:10px;box-shadow:0 6px 24px #0005;font:14px/1.5 "Microsoft JhengHei",sans-serif;overflow:hidden;text-align:left;display:flex;flex-direction:column;}
      #bcd-panel *{box-sizing:border-box;} #bcd-panel .body{padding:12px;min-height:0;flex:1;overflow:auto;overscroll-behavior:contain;} #bcd-panel .head{flex-shrink:0;}
      #bcd-panel h2{margin:0;font-size:16px;} #bcd-panel p{margin:8px 0;white-space:pre-line;overflow-wrap:anywhere;color:#c6d0df;} #bcd-panel .head{padding:10px 12px;background:#2b3242;display:flex;align-items:center;justify-content:space-between;}
      #bcd-panel button{font:inherit;color:#fff;background:#34506e;border:1px solid #5b7491;border-radius:5px;padding:8px;cursor:pointer;min-height:38px;white-space:normal;}
      #bcd-panel button:hover{background:#426286;} #bcd-panel button:disabled{opacity:.55;cursor:default;} #bcd-panel :focus-visible{outline:2px solid #b7dfff;outline-offset:2px;}
      #bcd-panel label{display:block;color:#c6d0df;margin:6px 0;} #bcd-panel input[type=text],#bcd-panel textarea,#bcd-panel select{display:block;font:inherit;color:#e8ecf3;background:#2b3242;border:1px solid #68788e;border-radius:4px;min-height:38px;padding:6px;width:100%;} #bcd-panel textarea{resize:vertical;min-height:78px;}
      #bcd-panel input[type=checkbox]{appearance:auto;width:18px;height:18px;margin:0 8px 0 0;vertical-align:middle;accent-color:#88b6dc;} #bcd-panel a{color:#acd5fc;text-decoration:underline;} #bcd-panel .actions,#bcd-panel .fields{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:8px 0;}
      .bcd-search-choice{box-sizing:border-box;display:inline-flex;align-items:center;align-self:flex-start;flex-shrink:0;gap:7px;width:fit-content;max-width:100%;min-height:34px;padding:5px 9px;margin:0 0 8px;background:#f4f7fa;color:#234764;border:1px solid #b5c5d3;border-radius:6px;font:13px/1.4 "Microsoft JhengHei",sans-serif;white-space:nowrap;cursor:pointer;text-decoration:none;}
      .bcd-search-choice:hover{background:#e8f0f7;border-color:#386589;} .bcd-search-choice[data-selected=true]{background:#e4eff8;border-color:#386589;color:#183d59;}
      .bcd-search-choice input[type=checkbox]{appearance:auto;width:16px;height:16px;flex:none;margin:0;accent-color:#386589;cursor:pointer;} .bcd-search-choice:focus-within{outline:2px solid #386589;outline-offset:2px;}
      .bcd-card-shell{display:flex;flex-direction:column;min-width:0;} .bcd-card-shell>a.product-item{display:block;width:100%;min-width:100%;max-width:100%;flex-basis:auto;padding:0;}
      #bcd-panel .book{padding:10px 0;border-top:1px solid #465167;} #bcd-panel .book>label{display:flex;align-items:flex-start;color:#fff;overflow-wrap:anywhere;} #bcd-panel .book input[type=checkbox]{flex:none;margin-top:3px;} #bcd-panel summary{cursor:pointer;padding:6px 0;color:#acd5fc;} #bcd-panel .warning{color:#ffc2b7;} #bcd-panel [hidden]{display:none!important;} #bcd-panel .exports button{width:100%;margin-top:8px;} #bcd-panel #bcd-status{padding:10px 0;border-top:1px solid #465167;} #bcd-panel .step{margin:16px 0 6px;font-size:15px;} #bcd-panel .primary{width:100%;background:#386589;} #bcd-panel .advanced{margin-top:12px;border-top:1px solid #465167;}
      #bcd-panel .wizard-nav{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:6px;position:sticky;top:-12px;padding:10px 0;background:#1f2430;z-index:1;} #bcd-panel .wizard-nav button{background:transparent;padding:8px 2px;} #bcd-panel .wizard-nav button[aria-current=step]{background:#386589;border-color:#acd5fc;color:#fff;} #bcd-panel .wizard-footer{position:sticky;bottom:-12px;background:#1f2430;padding:8px 0 4px;display:flex;gap:8px;border-top:1px solid #465167;} #bcd-panel .wizard-footer button{flex:1;}
      #bcd-panel .combine-books{padding:8px 0;border-block:1px solid #465167;margin:12px 0;} #bcd-panel .combine-books>summary{font-weight:700;font-size:15px;} #bcd-panel .work-preview h4{font-size:15px;margin:16px 0 4px;overflow-wrap:anywhere;} #bcd-panel .work-preview ol{padding-left:24px;margin:8px 0;} #bcd-panel .work-preview li{padding:6px 0;white-space:pre-line;overflow-wrap:anywhere;} #bcd-panel .combine-review>label{display:flex;gap:8px;align-items:flex-start;color:#e8ecf3;} #bcd-panel .combine-review input[type=checkbox]{flex:none;margin-top:3px;}
      #bcd-panel.collapsed{width:auto;max-width:calc(100vw - 24px);} #bcd-panel .head{gap:12px;} #bcd-panel h2{overflow-wrap:anywhere;flex:1;min-width:0;} #bcd-panel .selection-badge{font-size:12px;color:#c6d0df;white-space:nowrap;font-variant-numeric:tabular-nums;} #bcd-panel .head button{padding:5px 10px;min-height:34px;flex:none;}
      #bcd-panel .actions{grid-template-columns:repeat(2,minmax(0,1fr));} #bcd-panel .actions button{min-width:0;} #bcd-panel .wizard-footer button:last-of-type{background:#386589;border-color:#88b6dc;} #bcd-panel .selection-help{font-size:12px;} #bcd-panel .book>label{margin:0 0 5px;font-weight:600;} #bcd-panel .book>p{font-size:12px;margin:5px 0;} #bcd-panel .volume-list>summary{font-weight:600;padding:10px 0;} #bcd-panel .volume-list>label{padding:4px 0;} #bcd-panel .body{scrollbar-color:#68788e #1f2430;} #bcd-panel ::selection{background:#386589;color:#fff;}
      @media(max-width:480px){#bcd-panel{right:12px;bottom:12px;font-size:16px;}} @media print{#bcd-panel{display:none;}}`;
    const head=el('div',null,panel); head.className='head'; el('h2',SITE_NAME+'工具箱',head);
    ui.badge=el('span','已选 0',head);ui.badge.className='selection-badge';
    const body=el('div',null,panel); body.className='body';body.id='bcd-body';body.hidden=true;panel.classList.add('collapsed');
    const collapse=button('展开',() => {body.hidden=!body.hidden;panel.classList.toggle('collapsed',body.hidden); collapse.textContent=body.hidden?'展开':'收起'; collapse.setAttribute('aria-expanded',String(!body.hidden));},head); collapse.setAttribute('aria-expanded','false');collapse.setAttribute('aria-controls','bcd-body');
    const nav=el('nav',null,body);nav.className='wizard-nav';nav.setAttribute('aria-label','导出步骤');
    const steps=[1,2,3].map(number=>{const section=el('section',null,body);section.dataset.step=number;return section;});
    function showStep(number) {
      state.step=number;steps.forEach((section,index)=>{section.hidden=index!==number-1;});
      nav.querySelectorAll('button').forEach((n,index)=>{if(index===number-1)n.setAttribute('aria-current','step');else n.removeAttribute('aria-current');});
      if(ui.back) {ui.back.hidden=number===1;ui.next.hidden=number===3;ui.next.textContent=number===1?'下一步：读取并核对':'下一步：导出';}
      body.scrollTop=0;
    }
    ui.showStep=showStep;
    ['1. 选卷册','2. 核对资料','3. 导出'].forEach((label,index)=>{button(label,()=>showStep(index+1),nav).dataset.idle='true';});
    const [selection,review,output]=steps;
    const rec=state.items[0]?.record;
    const visibleCover=originalCover(document.querySelector('#mainpic a.nbg')?.href || document.querySelector('meta[property="og:image"]')?.content || document.querySelector('img.cover')?.src);
    if(id && (rec?.cover || visibleCover)) button('下载本书封面原图',() => void downloadOriginal(rec?.cover ? rec : {cover:visibleCover,title:clean(document.querySelector('h1')?.textContent) || id}),selection).dataset.idle='true';
    if(id) {
      const series=relatedSeries(document,location.href);
      for(const [index,url] of series.entries()) link(series.length>1?`打开本站系列页 ${index+1}，核对整套卷册`:'打开本站系列页，选择整套卷册',url,selection);
      el('p',series.length?'当前单本已读取；可导出单本，或打开系列页选整套。':'当前单本已读取。其他册未列成套书时，可用下方“多册合成作品”提供链接。',selection);
      const embedded=CURRENT_SITE?.collect?.(document,location.href) || [];
      if(embedded.some(row=>productId(new URL(row.url,location.href).href)!==id)) button('载入本页明确套书目录',()=>{
        const items=collectItems(document);mergeItems(items);state.currentIds=new Set(items.map(i=>i.id));state.selectionChanged=true;
        progress(`已载入 ${items.length} 条本页套书目录，请选择所需卷册；不代表全套完整。`);render();
      },selection).dataset.idle='true';
    } else el('p',variants.length ? '本页包含多个书目／期数组合。展开卷册列表勾选所需版本与册号；采用该变体的封面和价格，缺失作者、书号与简介请补录。' : searchResults ? '搜索结果可能包含其他作品。先勾选所需卷册，再分析或读取；不读取推荐区，翻页后保留选择，缺少的册可补链接。' : CURRENT_SITE ? '只读取当前页面已显示的书目，不保证全套完整；不同作品 / 版本请取消勾选。可翻页后刷新本页书目或手动补卷链接。' : DOUBAN ? '默认选择已识别卷册（不含合集）。先读取完整目录，再核对选择。' : '默认选择当前系列的非合集卷册，请核对。',selection);
    if(DOUBAN && isSeries) {const readDirectory=button('读取 / 继续读取整个丛书目录',()=>void loadDirectory(),selection);readDirectory.className='primary';readDirectory.dataset.idle='true';}
    const combine=el('details',null,selection);combine.className='combine-books';combine.open=state.combine;
    el('summary','多册合成作品',combine);
    el('p','不用网站列成套书。勾选当前页卷册，或粘贴本站的多册商品链接，与已勾选卷册一起分析；之后核对作品名和卷号。',combine);
    const combineLabel=el('label','各册商品链接（每行一条）',combine);ui.combineLinks=el('textarea',null,combineLabel);
    ui.combineLinks.placeholder=CURRENT_SITE ? location.origin+'/…（本站商品链接）' : DOUBAN ? 'https://book.douban.com/subject/…/' : 'https://www.books.com.tw/products/…';
    const combineActions=el('div',null,combine);combineActions.className='actions';
    button('分析商品链接',()=>void addLinks(true),combineActions).dataset.idle='true';
    button('分析已勾选卷册',()=>{
      if(state.busy)return;
      if(selected().length<2){progress('请先勾选至少 2 册，或粘贴其他册的商品链接。');return;}
      state.combine=true;state.combineApproved='';showStep(2);void run('read');
    },combineActions).dataset.idle='true';
    button('返回普通导出',()=>{state.combine=false;state.combineApproved='';progress('已返回普通导出，卷册选择和资料保留。');render();},combine).dataset.idle='true';
    ui.count=el('p','',body);nav.after(ui.count);
    const listDetails=el('details',null,selection);listDetails.className='volume-list';listDetails.open=!searchResults;ui.listDetails=listDetails;
    if(variants.length)ui.variantBook=selectControl('书目 / 版本',[['','全部书目'],...[...new Set(variants.map(row=>row.book))].map(book=>[book,book])],'',listDetails,render);
    ui.listSummary=el('summary','查看 / 调整卷册选择与逐本资料',listDetails);
    const filterLabel=el('label',null,listDetails);ui.onlySelected=el('input',null,filterLabel);ui.onlySelected.type='checkbox';el('span','仅看已选卷册',filterLabel);ui.onlySelected.addEventListener('change',render);
    ui.list=el('div',null,listDetails);
    const readSelected=button('读取 / 继续读取所选资料',() => void run('read'),review);readSelected.className='primary';readSelected.dataset.idle='true';
    ui.stop=button('立即暂停',() => {state.stop=true; state.abort?.(); progress('正在暂停…');},body);
    ui.status=el('p','',body); ui.status.id='bcd-status'; ui.status.setAttribute('role','status'); ui.status.setAttribute('aria-live','polite');
    ui.blocked=link('打开待处理商品（登录 / 验证）','#',body);
    const settings=el('div',null,review);el('h3','作品与卷册资料（自动整理，可修改）',settings).className='step';
    ui.inference=el('p','',settings);
    ui.exportScope=selectControl('导出用途',[['auto','自动建议'],['volume','单卷'],['work','作品']],['auto','volume','work'].includes(state.restoredExportScope) ? state.restoredExportScope : 'auto',settings,saveDraft);
    ui.exportHint=el('p','读取后根据已选书目建议用途，仍可手动选择。',settings);
    ui.grouping=selectControl('归组方式',[['auto','自动识别作品'],['same','按同一作品归组（我已核对）']],'auto',settings);
    const selectionOptions=el('details',null,selection);selectionOptions.className='selection-options';el('summary','批量选择 / 合集选项',selectionOptions);
    selection.prepend(selectionOptions);
    const advanced=el('details',null,review); advanced.className='advanced'; el('summary','高级设置 · 归组与版本',advanced);
    const actions=el('div',null,selection); actions.className='actions';selectionOptions.before(actions);
    button('全选本页',() => {state.selectionChanged=true;state.items.filter(i=>!isSeries || state.currentIds.has(i.id)).forEach(i=>{i.selected=ui.includeBundles.checked || !bundle(i);}); progress(`已选 ${selected().length} 本，请读取或核对资料后导出。`); render();},actions).dataset.idle='true';
    button('全部取消',() => {state.selectionChanged=true;state.items.forEach(i=>{i.selected=false;}); progress('已取消全部勾选，不会导出未选择的书。'); render();},actions).dataset.idle='true';
    const includeLabel=el('label',null,selectionOptions); ui.includeBundles=el('input',null,includeLabel); ui.includeBundles.type='checkbox';
    el('span','全选时包含合集／盒装／合售（仅供核对）',includeLabel);
    const readActions=el('div',null,selectionOptions); readActions.className='actions';
    if(isSeries && !DOUBAN && !variants.length) {
      const paging=el('div',null,selection);paging.className='actions';actions.after(paging);
      ui.nextListing=button(CURRENT_SITE?.key==='eslite'?'前往下一页 · 保留已选':'加载下一页书目',()=>{
        if(CURRENT_SITE?.key==='eslite') {
          const active=document.querySelector('.search-products-pagination [data-gid="pagination-page"].active'), pages=[...document.querySelectorAll('.search-products-pagination [data-gid="pagination-page"][data-position]')];
          const next=document.querySelector('.search-products-pagination [data-gid="pagination-next"] button');
          if(!next || next.disabled || active && pages.length && Number(active.dataset.position)>=Math.max(...pages.map(n=>Number(n.dataset.position)))) {progress('当前页面没有可用的下一页，请使用网站分页。');return;}
          saveDraft();next.click();
        } else void loadNextListingPage();
      },paging);ui.nextListing.dataset.idle='true';
      el('p','翻页保留已选；加载下一页只读书目，新增项需自行勾选。',selection).className='selection-help';
    }
    if(DOUBAN && isSeries) {
      button('全选已载入卷册',()=>{selectLoaded();progress(`已选 ${selected().length} 本；合集默认不选，可逐项修改。`);render();},readActions).dataset.idle='true';
      button('全选卷册并继续读取',()=>{
        if(state.busy) return;
        selectLoaded(); void run('read');
      },readActions).dataset.idle='true';
      el('p','全选包含其他分页已载入卷册；合集按上方开关处理。',selectionOptions);
    }
    if(isSeries) button('刷新本页书目',() => {
      const items=variants.length ? citiVariants(document,location.href).map(row=>({id:productId(row.url),title:row.title,selected:false})) : collectItems(document); state.currentIds=new Set(items.map(i=>i.id));
      try {
        if(DOUBAN) {
          const sid=location.pathname.match(/\d+/)[0], parsed=seriesPage(document,location.href,sid);
          const dir=state.directory || {id:sid,done:[],pending:[]};
          dir.done=[...new Set([...dir.done,parsed.url])];
          dir.pending=[...new Set([...dir.pending,...parsed.links])].filter(u=>!dir.done.includes(u));
          dir.complete=!dir.pending.length; state.directory=dir;
        }
        mergeItems(items,DOUBAN);progress('已合并当前页书目，其他页的勾选与修改保留。');
      } catch(err) {if(state.directory) state.directory.complete=false;progress(err.message);}
      render();
    },isSeries && !DOUBAN && !variants.length ? ui.nextListing.parentElement : readActions).dataset.idle='true';
    ui.work=inputControl('统一作品名', '',settings); ui.work.maxLength=240; ui.work.placeholder='留空保留各卷作品名';
    ui.authors=inputControl('统一作者','',settings); ui.authors.maxLength=500;ui.authors.placeholder='读取后自动填写，冲突时保留各卷';
    ui.publisher=inputControl('统一出版社','',settings); ui.publisher.maxLength=160;
    const group=el('div',null,advanced); group.className='fields';
    if(DOUBAN || CURRENT_SITE) ui.region=selectControl('统一出版地区',[['','保留各卷地区'],...REGIONS],'',group);
    ui.edition=selectControl('统一版本',[['','保留各卷版本'],...EDITIONS],'',group);
    ui.content=selectControl('统一内容类型',[['','保留各卷类型'],...CONTENTS],'',group);
    const setFields=el('div',null,settings); setFields.className='fields';
    ui.sequence=selectControl('卷册形式',[['','保留各卷形式'],...FORMATS],'',setFields);
    ui.total=inputControl('总卷数 / 已出版至','',setFields); ui.total.inputMode='numeric'; ui.total.placeholder='不确定留空';
    ui.serial=selectControl('连载状态',[['','保留各卷状态'],...SERIAL],'',setFields);
    for(const key of ['grouping','work','authors','publisher','region','edition','sequence','total','serial','content']) if(ui[key]) {
      ui[key].addEventListener(ui[key].tagName==='SELECT'?'change':'input',()=>{state.dirty.add(key); refreshAuto(); saveDraft();});
      if(Object.hasOwn(state.restoredManual || {},key)) {ui[key].value=state.restoredManual[key];state.dirty.add(key);}
    }
    button('重新自动整理（清除整套手填项）',()=>{state.dirty.clear();ui.grouping.value='auto';render();},advanced).dataset.idle='true';
    el('p','总卷数只填正篇，不含公式书、外传。数量不按本页册数推算。',settings);
    ui.gaps=el('p','',review);ui.gaps.className='warning';
    ui.combineReview=el('section',null,review);ui.combineReview.className='combine-review';
    el('h3','作品与卷册预览',ui.combineReview).className='step';ui.combineSummary=el('p','',ui.combineReview);
    ui.workPreview=el('div',null,ui.combineReview);ui.workPreview.className='work-preview';
    el('p','卷号待确认时，回到第 1 步打开“核对 / 修改本卷资料”修改；上下或上中下选择对应卷册形式。',ui.combineReview);
    const confirmLabel=el('label',null,ui.combineReview);ui.combineConfirm=el('input',null,confirmLabel);ui.combineConfirm.type='checkbox';
    el('span','已核对这些卷册属于同一作品和版本，确认按此导出',confirmLabel);
    ui.combineConfirm.addEventListener('change',()=>{state.combineApproved=ui.combineConfirm.checked ? state.combineSignature : '';});
    const missing=el('details',null,review);missing.className='missing-links';el('summary','目录少了卷册？粘贴链接补入',missing);
    const linksLabel=el('label','缺少卷册的商品链接（每行一条）',missing);ui.links=el('textarea',null,linksLabel);ui.links.placeholder=CURRENT_SITE ? location.origin+'/…（本站单本商品链接）' : DOUBAN?'https://book.douban.com/subject/…/':'https://www.books.com.tw/products/…';
    el('p','只补入单本，不覆盖原勾选或手填项。不同作品 / 版本会暂不勾选。',missing);
    button('读取并补入这些卷册',()=>void addLinks(),missing).dataset.idle='true';
    const exports=el('div',null,output);exports.className='exports';
    el('h3','这份资料用来做什么？',exports).className='step';
    button('资料用于搜书 · XLSX（不带图）',()=>void run('search'),exports).dataset.idle='true';
    el('p','到电脑“搜书 → 导入 ISBN 搜书清单”，继续联网查资料。',exports);
    button('带封面直接入库 · ZIP',()=>void run('local'),exports).dataset.idle='true';
    el('p','包含已核对资料与封面。到“我的书库 → 导入本地建档包”，预览确认后入库。',exports);
    el('p','不会自动写入书库，也不会改动现有藏书。',exports);
    const footer=el('div',null,body);footer.className='wizard-footer';
    ui.back=button('上一步',()=>showStep(state.step-1),footer);ui.back.dataset.idle='true';
    ui.next=button('',()=>{const previous=state.step;showStep(previous+1);if(previous===1 && selected().some(i=>!i.record || i.error))void run('read');},footer);ui.next.dataset.idle='true';
    footer.appendChild(ui.stop);
    showStep(1);
    if(!state.items.length) state.message='未找到当前页的系列书目。请确认已加载书目，或打开单本页面读取；不会扫描推荐区。';
    document.addEventListener('change',event => {
      if(state.busy || !event.target.matches?.('input.prodCheck')) return;
      state.selectionChanged=true;
      for(const item of state.items) {const check=[...document.querySelectorAll('input.prodCheck')].find(c=>c.value===item.id); if(check) item.selected=check.checked;}
      render();
    });
    render();
    const dynamicRows=searchResults && ['eslite','taaze'].includes(CURRENT_SITE?.key) ? SEARCH_CHOICES[CURRENT_SITE.key].rows : isSeries && CURRENT_SITE?.key==='chingwin' ? '.productBox' : isSeries && CURRENT_SITE?.key==='shogakukan' ? '#main .bookBox01 li.clearfix' : isSeries && CURRENT_SITE?.key==='kodansha' ? '[data-scroll-anchor="publications"] li' : '';
    if(dynamicRows) {
      // These search pages load or replace cards after document-idle/view changes.
      const selector=dynamicRows; let pending;
      const observer=new MutationObserver(changes=>{
        const relevant=changes.some(change=>{
          const target=change.target.nodeType===1 ? change.target : change.target.parentElement;
          if(target?.closest('#bcd-panel,.bcd-search-choice')) return false;
          const nodes=[...change.addedNodes,...change.removedNodes];
          if(nodes.length && nodes.every(node=>node.nodeType===1 && node.matches('.bcd-search-choice'))) return false;
          return target?.closest(selector) || nodes.some(node=>node.nodeType===1 && (node.matches(selector) || node.querySelector(selector)));
        });
        if(!relevant) return;
        clearTimeout(pending); pending=setTimeout(()=>{
          try {
            if(listingDirectoryId(location.href)!==state.directory.id) {
              observer.disconnect();saveDraft();
              document.querySelectorAll('.bcd-search-choice').forEach(n=>n.remove());document.querySelectorAll('.bcd-card-shell').forEach(n=>n.replaceWith(...n.children));panel.remove();
              Object.assign(state,{items:[],directory:null,dirty:new Set(),auto:new Map(),selectionChanged:false,combine:false,restoredManual:{},warning:'',blockedUrl:''});boot();return;
            }
            const items=collectItems(document); state.currentIds=new Set(items.map(item=>item.id)); mergeItems(items,false); render();
          } catch(err) {progress(err.message);}
        },100);
      });observer.observe(document.body,{childList:true,subtree:true,characterData:true});
    }
  }
  boot();
})();
