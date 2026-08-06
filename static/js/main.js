// ── Scroll reveal ─────────────────────────────────────────────────────────
const revealObs = new IntersectionObserver((entries) => {
  entries.forEach(entry => {
    if (entry.isIntersecting) {
      entry.target.classList.add('is-visible');
      revealObs.unobserve(entry.target); // animate once
    }
  });
}, { threshold: 0.12 });

document.querySelectorAll('.reveal').forEach(el => revealObs.observe(el));

// Stagger sibling .reveal elements inside cards/lists
document.querySelectorAll('.handbook-index, .announcement-list, .contents-grid, .chapter-list, .resource-grid')
  .forEach(container => {
    container.querySelectorAll('.reveal').forEach((el, i) => {
      el.style.transitionDelay = `${i * 0.07}s`;
    });
  });

// Stagger chapter sections on handbook section pages
document.querySelectorAll('.handbook-body .reveal').forEach((el, i) => {
  el.style.transitionDelay = `${i * 0.1}s`;
});

// ── Mobile hamburger menu ─────────────────────────────────────────────────
const burger = document.getElementById('hnav-burger');
const navLinks = document.querySelector('.hnav__links');
if (burger && navLinks) {
  burger.addEventListener('click', () => {
    const open = navLinks.classList.toggle('hnav__links--open');
    burger.setAttribute('aria-expanded', open);
    burger.innerHTML = open
      ? '<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M4 4l12 12M16 4L4 16" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>'
      : '<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M3 5h14M3 10h14M3 15h14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>';
  });
  // Close menu when a link is clicked
  navLinks.querySelectorAll('a').forEach(a => {
    a.addEventListener('click', () => {
      navLinks.classList.remove('hnav__links--open');
      burger.innerHTML = '<svg width="20" height="20" viewBox="0 0 20 20" fill="none"><path d="M3 5h14M3 10h14M3 15h14" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>';
    });
  });
}

// ── Announcement modal ────────────────────────────────────────────────────
const overlay    = document.getElementById('ann-overlay');
const modal      = document.getElementById('ann-modal');
const modalClose = document.getElementById('ann-modal-close');

if (overlay && modal) {
  function openModal(card) {
    const imgWrap = document.getElementById('ann-modal-img-wrap');
    const img     = document.getElementById('ann-modal-img');
    const imgSrc  = card.dataset.img || '';

    if (imgSrc) {
      img.src = imgSrc;
      imgWrap.classList.remove('ann-modal__img-wrap--hidden');
    } else {
      imgWrap.classList.add('ann-modal__img-wrap--hidden');
    }

    document.getElementById('ann-modal-tag').textContent   = card.dataset.tag   || '';
    document.getElementById('ann-modal-date').textContent  = card.dataset.date  || '';
    document.getElementById('ann-modal-title').textContent = card.dataset.title || '';
    document.getElementById('ann-modal-text').innerHTML    = card.dataset.body  || '';

    overlay.classList.add('ann-overlay--open');
    document.body.style.overflow = 'hidden';
  }

  function closeModal() {
    overlay.classList.remove('ann-overlay--open');
    document.body.style.overflow = '';
  }

  document.querySelectorAll('.ann-card').forEach(card => {
    card.addEventListener('click', () => openModal(card));
  });

  modalClose.addEventListener('click', e => { e.stopPropagation(); closeModal(); });
  overlay.addEventListener('click', e => { if (e.target === overlay) closeModal(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); });
}

// ── Auto-dismiss flash messages ───────────────────────────────────────────
document.querySelectorAll('.flash').forEach(f => {
  setTimeout(() => {
    f.style.transition = 'opacity 0.4s';
    f.style.opacity = '0';
    setTimeout(() => f.remove(), 400);
  }, 4000);
});
