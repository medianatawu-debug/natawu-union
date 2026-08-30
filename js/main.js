
(function(){
'use strict';
const body=document.body;
const toggle=document.querySelector('.mobile-toggle');
const nav=document.querySelector('.nav');
if(toggle&&nav){toggle.addEventListener('click',()=>{nav.classList.toggle('open');toggle.setAttribute('aria-expanded',nav.classList.contains('open'));});}
const moreBtn=document.querySelector('.more-btn'), moreMenu=document.querySelector('.more-menu');
if(moreBtn&&moreMenu){moreBtn.addEventListener('click',e=>{e.stopPropagation();moreMenu.classList.toggle('show')});document.addEventListener('click',()=>moreMenu.classList.remove('show'));}
const page=body.parentElement?.dataset.page || document.documentElement.dataset.page;
document.querySelectorAll('.nav a[data-page]').forEach(a=>{if(a.dataset.page===page)a.classList.add('active')});
// Gallery lightbox
const lightbox=document.getElementById('lightbox');
if(lightbox){const img=lightbox.querySelector('img');document.querySelectorAll('.gallery-card').forEach(card=>card.addEventListener('click',()=>{img.src=card.dataset.img;lightbox.classList.add('show')}));const close=()=>lightbox.classList.remove('show');document.getElementById('closeLightbox')?.addEventListener('click',close);lightbox.addEventListener('click',e=>{if(e.target===lightbox)close()});document.addEventListener('keydown',e=>{if(e.key==='Escape')close()});}
// Dynamic social SVG icons
const icons={facebook:'<svg viewBox="0 0 24 24" fill="currentColor"><path d="M14 8h3V4h-3c-3.3 0-5 2-5 5v3H6v4h3v4h4v-4h3.1l.9-4H13V9c0-.7.3-1 1-1Z"/></svg>',instagram:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="5"/><circle cx="12" cy="12" r="4"/><circle cx="17.5" cy="6.5" r="1" fill="currentColor" stroke="none"/></svg>',tiktok:'<svg viewBox="0 0 24 24" fill="currentColor"><path d="M15 3c.5 2.4 1.9 3.9 4 4.3v3.2c-1.5-.1-2.8-.6-4-1.4v7.1A5.8 5.8 0 1 1 9.2 10c.4 0 .8 0 1.2.1v3.3a2.7 2.7 0 1 0 1.5 2.4V3h3.1Z"/></svg>'};
const socialData=[['facebook','https://www.facebook.com/profile.php?id=61578191977447','Facebook'],['instagram','https://www.instagram.com/natawu_union','Instagram'],['tiktok','https://www.tiktok.com/@natawu_union','TikTok']];
document.querySelectorAll('.socials').forEach(box=>{box.innerHTML='';socialData.forEach(([key,url,label])=>{const a=document.createElement('a');a.href=url;a.target='_blank';a.rel='noopener noreferrer';a.setAttribute('aria-label',label);a.title=label;a.innerHTML=icons[key];box.appendChild(a)});});
// Contact form -> Google Apps Script. Replace URL after deployment.
const CONTACT_SCRIPT_URL='https://script.google.com/macros/s/AKfycbzqJ0F8esP885iibMern1kxSYNiz0z07gICcTrqMx3ksvAJ-fzy-NPFyMDoCwaN-sGK/exec';
const contact=document.getElementById('contactForm');
if(contact){contact.addEventListener('submit',async e=>{e.preventDefault();const btn=contact.querySelector('button');const old=btn.textContent;btn.disabled=true;btn.textContent='Sending...';const data=Object.fromEntries(new FormData(contact).entries());data.action='contact';try{if(CONTACT_SCRIPT_URL.includes('PASTE_YOUR')){const email=data.branch||'sgift8083@gmail.com';window.location.href=`mailto:${email}?subject=${encodeURIComponent('NATAWU Website Enquiry - '+data.firstName+' '+data.lastName)}&body=${encodeURIComponent(data.message+'\n\nPhone: '+data.phone+'\nEmail: '+data.email)}`;}else{const r=await fetch(CONTACT_SCRIPT_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(data)});const result=await r.json();if(!result.success)throw new Error(result.message||'Unable to send');alert('Thank you. Your message has been sent to NATAWU.');contact.reset();}}catch(err){alert('The message could not be sent. Please try again or email NATAWU directly.');}finally{btn.disabled=false;btn.textContent=old;}});}
})();
