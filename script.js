const docxDocuments = window.DOCX_DOCUMENTS || {};
// EDIT HERE: Change this value to control how long the nav waits before collapsing.
const NAV_IDLE_DELAY_MS = 1000;

// EDIT HERE: Paste future public website URLs into externalUrl.
// Leave externalUrl as "" until you have a deployed site, Google Docs, Drive, Canva, Notion, or other web page.
const projects = [
  {
    id: "project1",
    number: "01",
    icon: "folder",
    title: "Quản lý tệp và thư mục",
    subtitle: "Thao tác cơ bản với tệp tin, thư mục và cấu trúc lưu trữ.",
    objective:
      "Tổ chức dữ liệu học tập theo cách khoa học, dễ tìm kiếm, dễ mở rộng và có minh chứng thao tác rõ ràng.",
    tools: ["File Explorer", "Windows", "OneDrive"],
    process: [
      "Tạo thư mục thực hành theo tên sinh viên.",
      "Tạo, đổi tên, sao chép và di chuyển tệp văn bản.",
      "Tạo thư mục con để phân loại tài liệu.",
      "Xóa, khôi phục và ghi lại ảnh minh chứng từng bước.",
    ],
    reflection:
      "Bài tập giúp hình thành thói quen quản lý dữ liệu có cấu trúc. Khi tên tệp và thư mục nhất quán, việc tìm kiếm và tái sử dụng tài liệu trở nên nhanh hơn.",
    externalUrl: "https://docs.google.com/document/d/1c7t7qoq1Hr_O1CFVrGFM3_PweIRiCoEm/edit?usp=drive_link&ouid=116043194491971042384&rtpof=true&sd=true",
    document: { type: "docx", source: "project1" },
  },
  {
    id: "project2",
    number: "02",
    icon: "search",
    title: "Tìm kiếm và đánh giá thông tin học thuật",
    subtitle: "Ứng dụng học sâu trong chẩn đoán hình ảnh y tế.",
    objective:
      "Biết cách tìm nguồn học thuật đáng tin cậy, so sánh độ tin cậy và trình bày chiến lược tìm kiếm có kiểm chứng.",
    tools: ["Google Scholar", "PubMed", "IEEE Xplore", "ScienceDirect"],
    process: [
      "Xác định phạm vi chủ đề và từ khóa chính.",
      "Dùng nhiều cơ sở dữ liệu học thuật để tìm nguồn.",
      "Đánh giá nguồn theo tác giả, nơi xuất bản, trích dẫn và tính cập nhật.",
      "Lập bảng tài liệu tham khảo và kết luận nguồn đáng tin.",
    ],
    reflection:
      "Tìm kiếm tốt không dừng ở việc có nhiều kết quả. Giá trị của bài nằm ở cách lọc, kiểm tra nguồn và giải thích vì sao một tài liệu đủ đáng tin để sử dụng.",
    externalUrl: "https://docs.google.com/document/d/1KhXw3wZ1VNflKLmDucIwnN7Dw2Vqt6nx/edit?usp=drive_link&ouid=116043194491971042384&rtpof=true&sd=true",
    document: { type: "docx", source: "project2" },
  },
  {
    id: "project3",
    number: "03",
    icon: "prompt",
    title: "Viết prompt hiệu quả",
    subtitle: "Thiết kế prompt phục vụ học tập với chủ đề Pluripotential Theory.",
    objective:
      "So sánh prompt cơ bản, cải tiến và nâng cao để thấy vai trò của bối cảnh, định dạng đầu ra và tiêu chí đánh giá.",
    tools: ["ChatGPT", "Prompt Engineering", "PDF report"],
    process: [
      "Chọn ba tác vụ học tập liên quan cùng một chủ đề khó.",
      "Viết nhiều cấp độ prompt cho từng tác vụ.",
      "So sánh đầu ra AI theo độ rõ, độ sâu và tính hữu ích.",
      "Rút ra nguyên tắc viết prompt cho học tập chuyên sâu.",
    ],
    reflection:
      "AI phản hồi tốt hơn khi người dùng cung cấp vai trò, bối cảnh, mục tiêu, ràng buộc và tiêu chí chất lượng rõ ràng.",
    externalUrl: "https://drive.google.com/file/d/1Brv0dquPnqUg7sBHbVXwFqBVMoLXFNgw/view?usp=drive_link",
    document: {
      type: "pdf",
      title: "Bài tập 3 - Thiết kế prompt phục vụ học tập",
      src: "assets/resources/project-3-prompt-design.pdf",
    },
  },
  {
    id: "project4",
    number: "04",
    icon: "network",
    title: "Hợp tác trực tuyến",
    subtitle: "Quản lý dự án nhóm bằng hệ sinh thái công cụ cộng tác.",
    objective:
      "Trình bày cách dùng công cụ số để lập kế hoạch, phân công, lưu trữ, giao tiếp và theo dõi tiến độ nhóm.",
    tools: ["ClickUp", "Google Drive", "Google Docs", "Messenger/Discord"],
    process: [
      "Thiết lập không gian quản lý dự án và tài khoản minh chứng.",
      "Tạo nhiệm vụ, nhãn, mức ưu tiên và phân công thành viên.",
      "Tổ chức kho tài liệu dùng chung và quyền truy cập.",
      "Ghi lại minh chứng về tương tác, tiến độ và vai trò cá nhân.",
    ],
    reflection:
      "Công cụ cộng tác chỉ hiệu quả khi đi kèm quy trình rõ ràng. Vai trò trưởng nhóm cần vừa tổ chức tài nguyên, vừa duy trì nhịp trao đổi và chất lượng nội dung.",
    externalUrl: "https://drive.google.com/file/d/17q0nXCiVKTKeJIsc3g8O-UiE9uR5-f4z/view?usp=drive_link",
    document: {
      type: "pdf",
      title: "Bài tập 4 - Áp dụng công cụ hợp tác trực tuyến",
      src: "assets/resources/project-4-online-collaboration.pdf",
    },
  },
  {
    id: "project5",
    number: "05",
    icon: "spark",
    title: "Sáng tạo nội dung với AI",
    subtitle: "Infographic Đạo hàm riêng và Nhân tử Lagrange.",
    objective:
      "Dùng AI để hỗ trợ xây dựng nội dung học thuật trực quan, dễ theo dõi và có quy trình sáng tạo rõ ràng.",
    tools: ["ChatGPT", "NotebookLM", "Canva AI", "Infographic"],
    process: [
      "Chuyển kiến thức toán học thành khung nội dung infographic.",
      "Thử nhiều cấp độ prompt để cải thiện chất lượng đầu ra.",
      "Dùng công cụ AI hỗ trợ minh họa và bố cục.",
      "Hoàn thiện sản phẩm cuối và phản tư vai trò của AI.",
    ],
    reflection:
      "AI hữu ích nhất khi được dùng như cộng sự tạo nháp và kiểm tra cấu trúc. Người học vẫn cần quyết định nội dung, tính chính xác và cách trình bày cuối cùng.",
    externalUrl: "https://drive.google.com/file/d/1zMTI07sBT5nQXBWP753A2BRgifdIW9-d/view?usp=drive_link",
    featuredImage: {
      src: "assets/resources/project-5-infographic.png",
      alt: "Infographic Đạo hàm riêng và Nhân tử Lagrange",
    },
    document: {
      type: "pdf",
      title: "Bài tập 5 - Ứng dụng AI sáng tạo nội dung",
      src: "assets/resources/project-5-ai-content.pdf",
    },
  },
  {
    id: "project6",
    number: "06",
    icon: "shield",
    title: "Sử dụng AI có trách nhiệm",
    subtitle: "Nguyên tắc cá nhân và phân tích chính sách học thuật.",
    objective:
      "Xây dựng bộ nguyên tắc dùng AI minh bạch, có kiểm chứng, tôn trọng liêm chính học thuật và bảo vệ dữ liệu.",
    tools: ["AI research assistant", "UET/VNU sources", "Policy analysis"],
    process: [
      "Tìm hiểu chính sách và định hướng của UET/ĐHQGHN.",
      "Đối chiếu với hướng dẫn của các trường và tổ chức quốc tế.",
      "Phân tích các vấn đề đạo đức khi dùng AI.",
      "Viết bộ nguyên tắc cá nhân và infographic minh họa.",
    ],
    reflection:
      "Dùng AI có trách nhiệm nghĩa là giữ vai trò chủ thể của người học: minh bạch khi sử dụng, kiểm chứng đầu ra và tự chịu trách nhiệm với lập luận cuối cùng.",
    externalUrl: "https://docs.google.com/document/d/13hzweIrjMpEtnLosv5Y6VnK9mQLQz1jo/edit?usp=drive_link&ouid=116043194491971042384&rtpof=true&sd=true",
    document: { type: "docx", source: "project6" },
  },
];

const qs = (selector, root = document) => root.querySelector(selector);
const qsa = (selector, root = document) => [...root.querySelectorAll(selector)];

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function renderExternalButton(project) {
  if (!project.externalUrl) {
    return `
      <span class="button disabled" aria-disabled="true">Mở tài liệu gốc</span>
      <p class="url-placeholder">Chưa có URL. Dán website vào <strong>externalUrl</strong> của ${project.id} trong <strong>script.js</strong>.</p>
    `;
  }

  return `
    <a class="button primary" href="${escapeHtml(project.externalUrl)}" target="_blank" rel="noopener noreferrer">
      Mở tài liệu gốc
    </a>
  `;
}

function renderProjectIcon(icon) {
  const icons = {
    folder: `
      <svg viewBox="0 0 48 48" role="img" aria-label="Folder icon">
        <path class="icon-fill" d="M7 15.5c0-2.2 1.8-4 4-4h9.5l3.6 4H37c2.2 0 4 1.8 4 4v14c0 2.2-1.8 4-4 4H11c-2.2 0-4-1.8-4-4z"/>
        <path class="icon-line" d="M7 20h34M11 11.5h9.5l3.6 4H37c2.2 0 4 1.8 4 4v14c0 2.2-1.8 4-4 4H11c-2.2 0-4-1.8-4-4v-18c0-2.2 1.8-4 4-4z"/>
        <path class="icon-glint" d="M31 25l3-3 3 3-3 3z"/>
      </svg>`,
    search: `
      <svg viewBox="0 0 48 48" role="img" aria-label="Search icon">
        <circle class="icon-fill" cx="21" cy="21" r="11"/>
        <circle class="icon-line" cx="21" cy="21" r="11"/>
        <path class="icon-line" d="M29.5 29.5L39 39"/>
        <path class="icon-glint" d="M17 15h8"/>
      </svg>`,
    prompt: `
      <svg viewBox="0 0 48 48" role="img" aria-label="Prompt icon">
        <path class="icon-fill" d="M9 12h30v20H22l-8 6v-6H9z"/>
        <path class="icon-line" d="M9 12h30v20H22l-8 6v-6H9z"/>
        <path class="icon-line" d="M16 20h16M16 25h10"/>
        <path class="icon-glint" d="M32 17l2-2 2 2-2 2z"/>
      </svg>`,
    network: `
      <svg viewBox="0 0 48 48" role="img" aria-label="Collaboration icon">
        <path class="icon-line" d="M15 17l18 4M18 32l15-11M18 32l14 2"/>
        <circle class="icon-fill" cx="14" cy="16" r="6"/>
        <circle class="icon-fill" cx="35" cy="21" r="6"/>
        <circle class="icon-fill" cx="15" cy="33" r="6"/>
        <circle class="icon-line" cx="14" cy="16" r="6"/>
        <circle class="icon-line" cx="35" cy="21" r="6"/>
        <circle class="icon-line" cx="15" cy="33" r="6"/>
      </svg>`,
    spark: `
      <svg viewBox="0 0 48 48" role="img" aria-label="Creative content icon">
        <path class="icon-fill" d="M10 14h23v20H10z"/>
        <path class="icon-line" d="M10 14h23v20H10zM14 29l6-6 5 5 3-3 5 5"/>
        <path class="icon-glint" d="M35 8l2.2 6.4L44 17l-6.8 2.6L35 26l-2.2-6.4L26 17l6.8-2.6z"/>
      </svg>`,
    shield: `
      <svg viewBox="0 0 48 48" role="img" aria-label="Responsible AI icon">
        <path class="icon-fill" d="M24 7l15 6v10.5c0 8.5-5.7 14.1-15 17.5-9.3-3.4-15-9-15-17.5V13z"/>
        <path class="icon-line" d="M24 7l15 6v10.5c0 8.5-5.7 14.1-15 17.5-9.3-3.4-15-9-15-17.5V13z"/>
        <path class="icon-line" d="M17 24l5 5 10-11"/>
      </svg>`,
  };

  return `<span class="project-icon icon-${icon}" aria-hidden="true">${icons[icon] || icons.folder}</span>`;
}

function renderDocument(project) {
  const documentInfo = project.document;

  if (project.featuredImage) {
    return `
      <div class="featured-image">
        <img src="${project.featuredImage.src}" alt="${escapeHtml(project.featuredImage.alt)}" loading="lazy">
      </div>
      ${renderDocumentFrame(documentInfo)}
    `;
  }

  return renderDocumentFrame(documentInfo);
}

function renderDocumentFrame(documentInfo) {
  if (documentInfo.type === "pdf") {
    return `
      <div class="document-frame" role="region" aria-label="${escapeHtml(documentInfo.title)}">
        <object data="${documentInfo.src}" type="application/pdf">
          <iframe title="${escapeHtml(documentInfo.title)}" src="${documentInfo.src}"></iframe>
          <p>Trình duyệt hiện tại không hiển thị được PDF nhúng.</p>
        </object>
      </div>
    `;
  }

  const doc = docxDocuments[documentInfo.source];
  return `
    <article class="docx-article" aria-label="${escapeHtml(doc?.title || "Tài liệu DOCX")}">
      ${doc?.html || "<p>Không tìm thấy nội dung DOCX đã chuyển đổi.</p>"}
    </article>
  `;
}

function renderProjectCards() {
  const grid = qs("#project-grid");
  grid.innerHTML = projects
    .map(
      (project) => `
        <article class="project-card reveal">
          <div class="project-card-top">
            ${renderProjectIcon(project.icon)}
            <span class="project-index">${project.number}</span>
          </div>
          <h3>${project.title}</h3>
          <p>${project.subtitle}</p>
          <ul class="tag-list">
            ${project.tools.slice(0, 3).map((tool) => `<li>${tool}</li>`).join("")}
          </ul>
          <footer>
            <a class="button ghost" href="#${project.id}">Xem chi tiết</a>
          </footer>
        </article>
      `,
    )
    .join("");
}

function renderProjectDetails() {
  const details = qs("#project-details");
  details.innerHTML = projects
    .map(
      (project) => `
        <article class="project-detail reveal" id="${project.id}">
          <div class="detail-topline">
            <div>
              <p class="eyebrow">Bài ${project.number}</p>
              <h2>${project.title}</h2>
            </div>
            <div>${renderExternalButton(project)}</div>
          </div>
          <div class="detail-layout">
            <div>
              <h3>Mục tiêu</h3>
              <p>${project.objective}</p>
              <h3>Công cụ sử dụng</h3>
              <ul class="chip-list">
                ${project.tools.map((tool) => `<li>${tool}</li>`).join("")}
              </ul>
              <h3>Quy trình thực hiện</h3>
              <ol class="process-list">
                ${project.process.map((step) => `<li>${step}</li>`).join("")}
              </ol>
              <div class="reflection">
                <h3>Nhận xét cá nhân</h3>
                <p>${project.reflection}</p>
              </div>
            </div>
            <div>
              <h3>Tài liệu đầy đủ / Minh chứng</h3>
              ${renderDocument(project)}
            </div>
          </div>
        </article>
      `,
    )
    .join("");
}

function renderDocumentLibrary() {
  const library = qs("#document-library");
  library.innerHTML = projects
    .map(
      (project, index) => `
        <details class="document-item reveal" data-document-project="${project.id}" ${index === 0 ? "open" : ""}>
          <summary>Bài ${project.number}: ${project.title}</summary>
          <div class="document-inner"></div>
        </details>
      `,
    )
    .join("");

  qsa(".document-item").forEach((item) => {
    const hydrate = () => {
      const inner = qs(".document-inner", item);
      if (inner.dataset.loaded === "true") return;
      const project = projects.find((entry) => entry.id === item.dataset.documentProject);
      inner.innerHTML = renderDocument(project);
      inner.dataset.loaded = "true";
    };

    if (item.open) hydrate();
    item.addEventListener("toggle", () => {
      if (item.open) hydrate();
    });
  });
}

function setupNavigation() {
  const header = qs(".site-header");
  const toggle = qs(".nav-toggle");
  const nav = qs("#main-nav");
  let idleTimer;

  const shouldUseIdleCollapse = () =>
    window.matchMedia("(min-width: 761px)").matches &&
    !window.matchMedia("(prefers-reduced-motion: reduce)").matches;

  const expandNav = () => {
    header.classList.remove("is-idle");
    window.clearTimeout(idleTimer);
  };

  const scheduleIdle = () => {
    window.clearTimeout(idleTimer);
    if (!shouldUseIdleCollapse()) return;
    idleTimer = window.setTimeout(() => {
      if (!header.matches(":hover") && !header.contains(document.activeElement)) {
        header.classList.add("is-idle");
      }
    }, NAV_IDLE_DELAY_MS);
  };

  toggle.addEventListener("click", () => {
    expandNav();
    const isOpen = nav.classList.toggle("is-open");
    toggle.setAttribute("aria-expanded", String(isOpen));
  });

  header.addEventListener("pointerenter", expandNav);
  header.addEventListener("pointerleave", scheduleIdle);
  header.addEventListener("focusin", expandNav);
  header.addEventListener("focusout", scheduleIdle);
  header.addEventListener("pointerdown", expandNav);
  window.addEventListener("resize", () => {
    if (!shouldUseIdleCollapse()) header.classList.remove("is-idle");
    scheduleIdle();
  });

  qsa(".main-nav a").forEach((link) => {
    link.addEventListener("click", () => {
      expandNav();
      nav.classList.remove("is-open");
      toggle.setAttribute("aria-expanded", "false");
      scheduleIdle();
    });
  });

  scheduleIdle();
}

function setupProfileFallback() {
  const image = qs(".profile-photo img");
  image.addEventListener("error", () => {
    image.classList.add("is-missing");
  });
}

function setupRevealAnimations() {
  const targets = qsa(".reveal");
  if (!("IntersectionObserver" in window)) {
    targets.forEach((target) => target.classList.add("is-visible"));
    return;
  }

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-visible");
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.12 },
  );

  targets.forEach((target) => observer.observe(target));
}

function setupActiveNav() {
  const header = qs(".site-header");
  // EDIT HERE: Increase/decrease this if the active nav changes too early or too late.
  const NAV_ACTIVE_OFFSET_PX = 120;
  const links = qsa(".main-nav a");
  const sectionMap = [
    { element: qs("#intro"), navId: "intro" },
    { element: qs('[aria-labelledby="about-title"]'), navId: "intro" },
    { element: qs("#projects"), navId: "projects" },
    { element: qs("#documents"), navId: "documents" },
    { element: qs("#summary"), navId: "summary" },
  ].filter((entry) => entry.element);
  let ticking = false;

  header.dataset.activeSection = "intro";

  const getActiveSectionId = () => {
    const headerHeight = header.getBoundingClientRect().height;
    const scrollLine = window.scrollY + headerHeight + NAV_ACTIVE_OFFSET_PX;
    let activeId = "intro";

    for (const entry of sectionMap) {
      const sectionTop = entry.element.getBoundingClientRect().top + window.scrollY;
      if (scrollLine >= sectionTop) activeId = entry.navId;
    }

    return activeId;
  };

  const applyActiveSection = (activeId = getActiveSectionId()) => {
    header.dataset.activeSection = activeId;
    links.forEach((link) => {
      link.classList.toggle("active", link.getAttribute("href") === `#${activeId}`);
    });
  };

  const requestActiveUpdate = () => {
    if (ticking) return;
    ticking = true;
    window.requestAnimationFrame(() => {
      applyActiveSection();
      ticking = false;
    });
  };

  links.forEach((link) => {
    link.addEventListener("click", () => {
      const targetId = link.getAttribute("href")?.replace("#", "");
      if (targetId) applyActiveSection(targetId);
      window.setTimeout(requestActiveUpdate, 120);
    });
  });

  window.addEventListener("scroll", requestActiveUpdate, { passive: true });
  window.addEventListener("resize", requestActiveUpdate);
  applyActiveSection();
}

function setupMusicControl() {
  const audio = qs("#portfolio-audio");
  const button = qs("#music-toggle");
  if (!audio || !button) return;

  const label = qs(".music-label", button);
  const audioSrc = audio.getAttribute("src");
  let hasAudioFile = true;

  const setState = (state) => {
    button.classList.toggle("is-playing", state === "playing");
    button.classList.toggle("is-missing", state === "missing");
    button.setAttribute("aria-pressed", state === "playing" ? "true" : "false");

    if (state === "missing") {
      label.textContent = "Add music";
      button.title = "Place your file at assets/audio/portfolio-music.mp3";
    } else if (state === "playing") {
      label.textContent = "Mute";
      button.title = "Pause portfolio music";
    } else {
      label.textContent = "Play music";
      button.title = "Play looping portfolio music";
    }
  };

  const markMissing = () => {
    hasAudioFile = false;
    audio.pause();
    setState("missing");
  };

  audio.addEventListener("error", markMissing);
  audio.addEventListener("pause", () => {
    if (hasAudioFile) setState("paused");
  });
  audio.addEventListener("play", () => {
    if (hasAudioFile) setState("playing");
  });

  if (location.protocol === "http:" || location.protocol === "https:") {
    fetch(audioSrc, { method: "HEAD" })
      .then((response) => {
        if (!response.ok) markMissing();
      })
      .catch(() => {
        setState("paused");
      });
  }

  button.addEventListener("click", async () => {
    if (!hasAudioFile) {
      setState("missing");
      return;
    }

    if (audio.paused) {
      try {
        audio.volume = 0.55;
        await audio.play();
        localStorage.setItem("portfolioMusicPreference", "playing");
      } catch {
        markMissing();
      }
    } else {
      audio.pause();
      localStorage.setItem("portfolioMusicPreference", "paused");
    }
  });

  setState("paused");
}

function keepProjectsVisible() {
  qs("#projects")?.classList.add("is-visible");
  qsa("#projects .reveal").forEach((item) => item.classList.add("is-visible"));
}

renderProjectCards();
renderProjectDetails();
keepProjectsVisible();
renderDocumentLibrary();
setupNavigation();
setupProfileFallback();
setupRevealAnimations();
setupActiveNav();
setupMusicControl();
