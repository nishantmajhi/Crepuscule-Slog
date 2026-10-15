(function () {
  const c = document.createElement("canvas");
  c.width = c.height = 64;
  const ctx = c.getContext("2d");

  // Background: top half transparent → amber bottom half
  ctx.fillStyle = "#f5a623"; // var(--amber)
  ctx.fillRect(0, 35, 64, 29);

  // Border
  ctx.strokeStyle = "#1a1814"; // var(--ink)
  ctx.lineWidth = 4;
  ctx.strokeRect(2, 2, 60, 60);

  // Three horizontal lines (the crepuscule-slog-mark::after bars)
  ctx.strokeStyle = "rgba(26,24,20,0.5)";
  ctx.lineWidth = 3;
  [48, 40, 32].forEach((y) => {
    ctx.beginPath();
    ctx.moveTo(10, y);
    ctx.lineTo(54, y);
    ctx.stroke();
  });

  const link =
    document.querySelector("link[rel='icon']") ||
    document.createElement("link");
  link.rel = "icon";
  link.href = c.toDataURL();
  document.head.appendChild(link);
})();
