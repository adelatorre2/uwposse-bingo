/* ============================================================
   Posse Bingo — app logic (vanilla JS)
   Reads data/bingo_items.xlsx at runtime via SheetJS. No build step:
   edit the Excel file, push, and the site updates.

   Each phone gets its own randomly-shuffled 5x5 board per deck,
   persisted in localStorage. The middle square is always the deck's
   "connector" square (fixed position, matches the printed cards),
   the other 24 squares are shuffled independently per device.
   ============================================================ */

(function () {
  "use strict";

  /* --------- Config --------- */
  var DATA_URL = "data/bingo_items.xlsx";
  var SHEET_NAME = "Items";
  var STORAGE_PREFIX = "posseBingo:v1:";
  var CENTER_INDEX = 12; // middle of a 5x5 grid (0-indexed)
  var GRID_SIZE = 5;

  // All winning lines as arrays of cell indices (rows, columns, diagonals)
  var LINES = (function () {
    var lines = [];
    for (var r = 0; r < GRID_SIZE; r++) {
      var row = [];
      for (var c = 0; c < GRID_SIZE; c++) row.push(r * GRID_SIZE + c);
      lines.push(row);
    }
    for (var c2 = 0; c2 < GRID_SIZE; c2++) {
      var col = [];
      for (var r2 = 0; r2 < GRID_SIZE; r2++) col.push(r2 * GRID_SIZE + c2);
      lines.push(col);
    }
    var diag1 = [], diag2 = [];
    for (var i = 0; i < GRID_SIZE; i++) {
      diag1.push(i * GRID_SIZE + i);
      diag2.push(i * GRID_SIZE + (GRID_SIZE - 1 - i));
    }
    lines.push(diag1, diag2);
    return lines;
  })();

  /* --------- State --------- */
  var decks = {};          // deckName -> { items: [text...], center: text }
  var deckOrder = [];       // ordered deck names as they appear in the sheet
  var activeDeck = "Warm Up";
  var board = null;         // { texts: [25], marked: [25] }
  var celebratedLines = {}; // line signature -> true (per page load, avoids re-confetti on re-render)

  /* --------- DOM --------- */
  var $ = function (id) { return document.getElementById(id); };
  var statusEl, boardEl, progressEl, winBanner, winText, confettiLayer;

  /* ============================================================
     Loading + parsing
     ============================================================ */
  function loadItems() {
    var url = DATA_URL + "?v=" + Date.now();
    fetch(url, { cache: "no-store" })
      .then(function (res) {
        if (!res.ok) throw new Error("HTTP " + res.status);
        return res.arrayBuffer();
      })
      .then(function (buf) {
        var rows = parseWorkbook(buf);
        if (!rows.length) {
          showError(
            "No bingo squares found. Check that <code>data/bingo_items.xlsx</code> has a " +
            "sheet named <strong>Items</strong> with <em>deck</em>, <em>text</em>, and " +
            "<em>is_center</em> columns."
          );
          return;
        }
        buildDecks(rows);
        clearStatus();
        wireDeckButtons();
        setDeck(deckOrder.indexOf(activeDeck) === -1 ? deckOrder[0] : activeDeck);
      })
      .catch(function (err) {
        showError(
          "Couldn't load the bingo squares. Make sure " +
          "<code>data/bingo_items.xlsx</code> exists and try refreshing." +
          '<br><span class="err-detail">(' + escapeHtml(String(err.message || err)) + ")</span>"
        );
      });
  }

  function parseWorkbook(buf) {
    var out = [];
    var wb;
    try {
      wb = XLSX.read(new Uint8Array(buf), { type: "array" });
    } catch (e) {
      showError("The Excel file couldn't be read. It may be corrupted — re-save it and push again.");
      return out;
    }
    var sheet = wb.Sheets[SHEET_NAME] || wb.Sheets[wb.SheetNames[0]];
    if (!sheet) return out;

    var raw;
    try {
      raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });
    } catch (e) {
      return out;
    }

    for (var i = 0; i < raw.length; i++) {
      var rec = normalizeRow(raw[i]);
      if (rec) out.push(rec);
    }
    return out;
  }

  function pick(row, keys) {
    var map = {};
    for (var k in row) {
      if (Object.prototype.hasOwnProperty.call(row, k)) {
        map[String(k).trim().toLowerCase()] = row[k];
      }
    }
    for (var j = 0; j < keys.length; j++) {
      if (map[keys[j]] !== undefined) return map[keys[j]];
    }
    return "";
  }

  function normalizeRow(row) {
    var text = String(pick(row, ["text"]) || "").trim();
    if (!text) return null;
    var deck = String(pick(row, ["deck"]) || "").trim() || "Bingo";
    var isCenterRaw = String(pick(row, ["is_center"]) || "").trim().toLowerCase();
    var isCenter = isCenterRaw === "true" || isCenterRaw === "1" || isCenterRaw === "yes";
    return { deck: deck, text: text, isCenter: isCenter };
  }

  function buildDecks(rows) {
    decks = {};
    deckOrder = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!decks[r.deck]) {
        decks[r.deck] = { items: [], center: null };
        deckOrder.push(r.deck);
      }
      if (r.isCenter && !decks[r.deck].center) {
        decks[r.deck].center = r.text;
      } else {
        decks[r.deck].items.push(r.text);
      }
    }
    // Fallback: if a deck has no explicit center, use its first item.
    for (var d = 0; d < deckOrder.length; d++) {
      var name = deckOrder[d];
      if (!decks[name].center && decks[name].items.length) {
        decks[name].center = decks[name].items.shift();
      }
    }
  }

  /* ============================================================
     Board generation + persistence
     ============================================================ */
  function shuffled(arr) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var tmp = a[i]; a[i] = a[j]; a[j] = tmp;
    }
    return a;
  }

  function storageKey(deckName) {
    return STORAGE_PREFIX + deckName;
  }

  function newBoardFor(deckName) {
    var deck = decks[deckName];
    var pool = deck.items.slice();
    // If fewer than 24 items, allow repeats so the grid still fills.
    while (pool.length < 24) pool = pool.concat(deck.items);
    var chosen = shuffled(pool).slice(0, 24);

    var texts = new Array(25);
    var k = 0;
    for (var idx = 0; idx < 25; idx++) {
      if (idx === CENTER_INDEX) {
        texts[idx] = deck.center || "Free space";
      } else {
        texts[idx] = chosen[k++];
      }
    }
    return { texts: texts, marked: new Array(25).fill(false) };
  }

  function loadBoardFor(deckName) {
    try {
      var raw = window.localStorage.getItem(storageKey(deckName));
      if (raw) {
        var parsed = JSON.parse(raw);
        if (parsed && Array.isArray(parsed.texts) && parsed.texts.length === 25 &&
            Array.isArray(parsed.marked) && parsed.marked.length === 25) {
          return parsed;
        }
      }
    } catch (e) { /* fall through to a fresh board */ }
    return newBoardFor(deckName);
  }

  function saveBoard() {
    try {
      window.localStorage.setItem(storageKey(activeDeck), JSON.stringify(board));
    } catch (e) { /* localStorage unavailable — game still works this session */ }
  }

  /* ============================================================
     Deck switching
     ============================================================ */
  function wireDeckButtons() {
    var buttons = document.querySelectorAll(".mode-btn[data-deck]");
    buttons.forEach(function (btn) {
      // Hide deck buttons for decks that don't exist in the data.
      if (deckOrder.indexOf(btn.dataset.deck) === -1) {
        btn.style.display = "none";
        return;
      }
      btn.addEventListener("click", function () { setDeck(btn.dataset.deck); });
    });
  }

  function setDeck(deckName) {
    if (!decks[deckName]) return;
    activeDeck = deckName;
    board = loadBoardFor(activeDeck);
    celebratedLines = {};

    var buttons = document.querySelectorAll(".mode-btn[data-deck]");
    buttons.forEach(function (btn) {
      var isActive = btn.dataset.deck === activeDeck;
      btn.classList.toggle("is-active", isActive);
      btn.setAttribute("aria-selected", String(isActive));
    });

    hideWinBanner();
    renderBoard();
  }

  /* ============================================================
     Rendering
     ============================================================ */
  function renderBoard() {
    boardEl.innerHTML = "";
    for (var i = 0; i < 25; i++) {
      (function (idx) {
        var cell = document.createElement("button");
        cell.type = "button";
        cell.className = "cell";
        if (idx === CENTER_INDEX) cell.classList.add("is-center");
        if (board.marked[idx]) cell.classList.add("is-marked");
        cell.textContent = board.texts[idx];
        cell.setAttribute("aria-pressed", String(!!board.marked[idx]));
        cell.addEventListener("click", function () { toggleCell(idx); });
        boardEl.appendChild(cell);
      })(i);
    }
    renderProgress();
    highlightCompletedLines(false);
  }

  function renderProgress() {
    var markedCount = board.marked.filter(Boolean).length;
    var lines = completedLines(board.marked);
    var bits = [markedCount + " of 25 marked"];
    if (lines.length) {
      bits.push(lines.length + (lines.length === 1 ? " line" : " lines") + " complete");
    }
    progressEl.textContent = bits.join(" · ");
  }

  function highlightCompletedLines(animateNew) {
    var cells = boardEl.querySelectorAll(".cell");
    cells.forEach(function (c) { c.classList.remove("is-bingo-line"); });

    var lines = completedLines(board.marked);
    var newlyCompleted = [];
    var seenSignatures = {};

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      var sig = line.join(",");
      seenSignatures[sig] = true;
      for (var j = 0; j < line.length; j++) {
        cells[line[j]].classList.add("is-bingo-line");
      }
      if (!celebratedLines[sig]) {
        newlyCompleted.push(line);
      }
    }

    if (animateNew && newlyCompleted.length) {
      for (var k = 0; k < newlyCompleted.length; k++) {
        celebratedLines[newlyCompleted[k].join(",")] = true;
      }
      var blackout = board.marked.every(Boolean);
      celebrate(blackout);
    }
  }

  function completedLines(marked) {
    return LINES.filter(function (line) {
      return line.every(function (idx) { return marked[idx]; });
    });
  }

  /* ============================================================
     Interaction
     ============================================================ */
  function toggleCell(idx) {
    board.marked[idx] = !board.marked[idx];
    saveBoard();
    var cell = boardEl.children[idx];
    cell.classList.toggle("is-marked", board.marked[idx]);
    cell.setAttribute("aria-pressed", String(board.marked[idx]));
    renderProgress();
    highlightCompletedLines(true);
  }

  /* ============================================================
     Celebration
     ============================================================ */
  function celebrate(blackout) {
    showWinBanner(blackout ? "BLACKOUT! Full card!" : "BINGO!");
    burstConfetti();
  }

  function showWinBanner(text) {
    winText.textContent = text;
    winBanner.hidden = false;
    window.clearTimeout(showWinBanner._t);
    showWinBanner._t = window.setTimeout(hideWinBanner, 4200);
  }
  function hideWinBanner() {
    winBanner.hidden = true;
  }

  function burstConfetti() {
    var colors = ["#C5050C", "#9B0000", "#D4A017", "#ffffff", "#333333"];
    var count = 60;
    for (var i = 0; i < count; i++) {
      var piece = document.createElement("span");
      piece.className = "confetti-piece";
      piece.style.left = (Math.random() * 100) + "vw";
      piece.style.background = colors[Math.floor(Math.random() * colors.length)];
      var duration = 1.6 + Math.random() * 1.4;
      piece.style.animationDuration = duration + "s";
      piece.style.animationDelay = (Math.random() * 0.3) + "s";
      confettiLayer.appendChild(piece);
      (function (el, ms) {
        window.setTimeout(function () { el.remove(); }, ms * 1000 + 500);
      })(piece, duration + 0.3);
    }
  }

  /* ============================================================
     Board actions
     ============================================================ */
  function wireControls() {
    $("new-board-btn").addEventListener("click", function () {
      var hasMarks = board.marked.some(Boolean);
      if (hasMarks && !window.confirm("Start a brand-new shuffled board? This clears your current marks.")) {
        return;
      }
      board = newBoardFor(activeDeck);
      celebratedLines = {};
      saveBoard();
      hideWinBanner();
      renderBoard();
    });

    $("clear-btn").addEventListener("click", function () {
      if (!board.marked.some(Boolean)) return;
      if (!window.confirm("Clear all marks on this board? The squares stay the same.")) return;
      board.marked = new Array(25).fill(false);
      celebratedLines = {};
      saveBoard();
      hideWinBanner();
      renderBoard();
    });
  }

  /* ============================================================
     Status helpers
     ============================================================ */
  function showError(html) {
    statusEl.innerHTML = '<div class="msg">' + html + "</div>";
  }
  function clearStatus() { statusEl.innerHTML = ""; }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /* --------- Boot --------- */
  document.addEventListener("DOMContentLoaded", function () {
    statusEl = $("status");
    boardEl = $("board");
    progressEl = $("progress-hint");
    winBanner = $("win-banner");
    winText = $("win-text");
    confettiLayer = $("confetti-layer");

    wireControls();
    loadItems();
  });
})();
