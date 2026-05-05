// ナビゲーション スクロール時のスタイル変更
const nav = document.getElementById('nav');
const backToTop = document.getElementById('backToTop');

window.addEventListener('scroll', () => {
  const scrolled = window.scrollY > 60;
  nav.classList.toggle('scrolled', scrolled);
  backToTop.classList.toggle('visible', window.scrollY > 400);
});

// ハンバーガーメニュー
const hamburger = document.getElementById('hamburger');
const navMenu = document.getElementById('navMenu');

hamburger.addEventListener('click', () => {
  const isOpen = navMenu.classList.toggle('open');
  hamburger.setAttribute('aria-expanded', isOpen);
});

navMenu.querySelectorAll('a').forEach(link => {
  link.addEventListener('click', () => {
    navMenu.classList.remove('open');
    hamburger.setAttribute('aria-expanded', false);
  });
});

// FAQ アコーディオン
document.querySelectorAll('.faq__question').forEach(btn => {
  btn.addEventListener('click', () => {
    const isOpen = btn.getAttribute('aria-expanded') === 'true';
    // 他を全て閉じる
    document.querySelectorAll('.faq__question').forEach(b => {
      b.setAttribute('aria-expanded', 'false');
      b.nextElementSibling.style.display = 'none';
    });
    // クリックしたものを開閉
    if (!isOpen) {
      btn.setAttribute('aria-expanded', 'true');
      btn.nextElementSibling.style.display = 'block';
    }
  });
});

// スクロールアニメーション（Intersection Observer）
const observer = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      entry.target.style.opacity = '1';
      entry.target.style.transform = 'translateY(0)';
    }
  });
}, { threshold: 0.1 });

document.querySelectorAll('.feature, .service__card, .price__card, .faq__item, .voice__card').forEach(el => {
  el.style.opacity = '0';
  el.style.transform = 'translateY(24px)';
  el.style.transition = 'opacity 0.5s ease, transform 0.5s ease';
  observer.observe(el);
});

// 数字カウントアップアニメーション
function animateCount(el) {
  const target = parseInt(el.dataset.target, 10);
  const duration = 1800;
  const step = target / (duration / 16);
  let current = 0;
  const timer = setInterval(() => {
    current = Math.min(current + step, target);
    el.textContent = Math.floor(current).toLocaleString();
    if (current >= target) clearInterval(timer);
  }, 16);
}

const countObserver = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      animateCount(entry.target);
      countObserver.unobserve(entry.target);
    }
  });
}, { threshold: 0.5 });

document.querySelectorAll('.stat__count').forEach(el => countObserver.observe(el));

// お問い合わせフォーム送信（送信先は別途設定が必要）
document.getElementById('contactForm').addEventListener('submit', (e) => {
  e.preventDefault();
  // TODO: Formspree / Google Forms などに接続して実際に送信できるようにしてください
  alert('お問い合わせありがとうございます！\n2〜3営業日以内にご連絡いたします。');
  e.target.reset();
});
