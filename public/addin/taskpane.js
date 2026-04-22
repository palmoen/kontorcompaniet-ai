(() => {
  const state = {
    mailboxReady: false,
    currentItem: null,
    currentBodyText: "",
    activeTab: "reply",
    selectedAdjustment: "",
    lastReply: "",
    lastSummary: "",
    lastTasks: "",
    lastTools: "",
  };

  const els = {};

  Office.onReady(() => {
    bindElements();
    bindEvents();
    setStatus("Klar", "loading");
    loadCurrentEmailContext()
      .then(() => {
        setStatus("Aktiv e-post lastet", "loading", 1800);
      })
      .catch(() => {
        setStatus("Kunne ikke lese aktiv e-post", "error", 2400);
      });
  });

  function bindElements() {
    els.tabs = Array.from(document.querySelectorAll(".tab-btn"));
    els.panels = Array.from(document.querySelectorAll(".panel"));
    els.statusLine = document.getElementById("statusLine");
    els.statusText = document.getElementById("statusText");

    els.toneSelect = document.getElementById("toneSelect");

    els.generateReplyBtn = document.getElementById("generateReplyBtn");
    els.insertReplyBtn = document.getElementById("insertReplyBtn");
    els.copyReplyBtn = document.getElementById("copyReplyBtn");
    els.replyOutput = document.getElementById("replyOutput");
    els.regenerateBtn = document.getElementById("regenerateBtn");
    els.promptActions = Array.from(document.querySelectorAll(".prompt-action"));

    els.generateSummaryBtn = document.getElementById("generateSummaryBtn");
    els.summaryOutput = document.getElementById("summaryOutput");
    els.copySummaryBtn = document.getElementById("copySummaryBtn");

    els.generateTasksBtn = document.getElementById("generateTasksBtn");
    els.tasksOutput = document.getElementById("tasksOutput");
    els.copyTasksBtn = document.getElementById("copyTasksBtn");

    els.generateNewEmailBtn = document.getElementById("generateNewEmailBtn");
    els.generateComplaintBtn = document.getElementById("generateComplaintBtn");
    els.suggestToolsBtn = document.getElementById("suggestToolsBtn");
    els.toolPrompt = document.getElementById("toolPrompt");
    els.toolsOutput = document.getElementById("toolsOutput");
    els.copyToolsBtn = document.getElementById("copyToolsBtn");

    els.refreshContextBtn = document.getElementById("refreshContextBtn");
    els.clearAllBtn = document.getElementById("clearAllBtn");
  }

  function bindEvents() {
    els.tabs.forEach((tab) => {
      tab.addEventListener("click", () => {
        switchTab(tab.dataset.tab);
      });
    });

    els.promptActions.forEach((btn) => {
      btn.addEventListener("click", async () => {
        state.selectedAdjustment = btn.dataset.action || "";
        await generateReply(state.selectedAdjustment);
      });
    });

    els.regenerateBtn?.addEventListener("click", async () => {
      state.selectedAdjustment = "Lag en ny variant";
      await generateReply("Lag en ny variant");
    });

    els.generateReplyBtn?.addEventListener("click", async () => {
      state.selectedAdjustment = "";
      await generateReply();
    });

    els.insertReplyBtn?.addEventListener("click", async () => {
      await insertIntoMessage(els.replyOutput.value || "");
    });

    els.copyReplyBtn?.addEventListener("click", async () => {
      await copyText(els.replyOutput.value || "");
    });

    els.generateSummaryBtn?.addEventListener("click", async () => {
      await generateSummary();
    });

    els.copySummaryBtn?.addEventListener("click", async () => {
      await copyText(els.summaryOutput.value || "");
    });

    els.generateTasksBtn?.addEventListener("click", async () => {
      await generateTasks();
    });

    els.copyTasksBtn?.addEventListener("click", async () => {
      await copyText(els.tasksOutput.value || "");
    });

    els.generateNewEmailBtn?.addEventListener("click", async () => {
      switchTab("tools");
      await generateToolIntent("new_email");
    });

    els.generateComplaintBtn?.addEventListener("click", async () => {
      switchTab("tools");
      await generateToolIntent("complaint");
    });

    els.suggestToolsBtn?.addEventListener("click", async () => {
      switchTab("tools");
      await generateToolIntent("tools");
    });

    els.copyToolsBtn?.addEventListener("click", async () => {
      await copyText(els.toolsOutput.value || "");
    });

    els.refreshContextBtn?.addEventListener("click", async () => {
      await loadCurrentEmailContext();
      setStatus("Aktiv e-post oppdatert", "loading", 1800);
    });

    els.clearAllBtn?.addEventListener("click", () => {
      clearOutputs();
      setStatus("Felter tømt", "loading", 1400);
    });
  }

  function switchTab(tabName) {
    state.activeTab = tabName;

    els.tabs.forEach((tab) => {
      tab.classList.toggle("active", tab.dataset.tab === tabName);
    });

    els.panels.forEach((panel) => {
      panel.classList.toggle("active", panel.dataset.panel === tabName);
    });
  }

  async function loadCurrentEmailContext() {
    const item = Office.context?.mailbox?.item;
    state.currentItem = item || null;

    if (!item) {
      state.currentBodyText = "";
      return;
    }

    state.mailboxReady = true;

    const bodyText = await getBodyText(item);
    state.currentBodyText = [
      buildHeaderContext(item),
      bodyText ? `E-postinnhold:\n${bodyText}` : "",
    ]
      .filter(Boolean)
      .join("\n\n");
  }

  function buildHeaderContext(item) {
    const subject = item.subject || "";
    const from =
      item.from?.displayName || item.from?.emailAddress || "";
    const to = Array.isArray(item.to)
      ? item.to.map((r) => r.displayName || r.emailAddress).filter(Boolean).join(", ")
      : "";
    return [`Emne: ${subject}`, from ? `Fra: ${from}` : "", to ? `Til: ${to}` : ""]
      .filter(Boolean)
      .join("\n");
  }

  function getBodyText(item) {
    return new Promise((resolve) => {
      if (!item?.body?.getAsync) {
        resolve("");
        return;
      }

      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve((result.value || "").trim());
          return;
        }
        resolve("");
      });
    });
  }

  async function generateReply(adjustment = "") {
    try {
      ensureContextLoaded();
      setLoading(true, "Genererer svar...");

      const response = await callAi({
        intent: "reply",
        tone: els.toneSelect?.value || "standard",
        context: state.currentBodyText,
        instruction: adjustment
          ? `Juster eksisterende svar med følgende føring: ${adjustment}`
          : "Generer et kort, profesjonelt og uformelt norsk svarutkast.",
        previousDraft: adjustment ? (els.replyOutput.value || "") : "",
      });

      const text = extractText(response);
      state.lastReply = text;
      els.replyOutput.value = text;
      setStatus("Svar klart", "loading", 1600);
    } catch (error) {
      setStatus("Kunne ikke generere svar", "error", 2600);
    } finally {
      setLoading(false);
    }
  }

  async function generateSummary() {
    try {
      ensureContextLoaded();
      setLoading(true, "Oppsummerer tråd...");

      const response = await callAi({
        intent: "summary",
        tone: "standard",
        context: state.currentBodyText,
        instruction:
          "Oppsummer tråden på norsk i 3–5 korte bullets: hva saken gjelder, hva som er gjort, hva som mangler.",
      });

      const text = extractText(response);
      state.lastSummary = text;
      els.summaryOutput.value = text;
      setStatus("Oppsummering klar", "loading", 1600);
    } catch (error) {
      setStatus("Kunne ikke oppsummere", "error", 2600);
    } finally {
      setLoading(false);
    }
  }

  async function generateTasks() {
    try {
      ensureContextLoaded();
      setLoading(true, "Lager oppgaver...");

      const response = await callAi({
        intent: "tasks",
        tone: "standard",
        context: state.currentBodyText,
        instruction:
          "Foreslå konkrete oppgaver, avtaler og oppfølging på norsk. Kort og oversiktlig.",
      });

      const text = extractText(response);
      state.lastTasks = text;
      els.tasksOutput.value = text;
      setStatus("Oppgaver klare", "loading", 1600);
    } catch (error) {
      setStatus("Kunne ikke lage oppgaver", "error", 2600);
    } finally {
      setLoading(false);
    }
  }

  async function generateToolIntent(intent) {
    try {
      const promptText = (els.toolPrompt?.value || "").trim();
      const combinedContext = [state.currentBodyText, promptText ? `Brukerprompt:\n${promptText}` : ""]
        .filter(Boolean)
        .join("\n\n");

      if (!combinedContext) {
        setStatus("Legg inn prompt eller åpne en e-post først", "error", 2400);
        return;
      }

      const loadingText =
        intent === "new_email"
          ? "Lager ny e-post..."
          : intent === "complaint"
            ? "Lager reklamasjonsutkast..."
            : "Analyserer verktøy...";

      setLoading(true, loadingText);

      const response = await callAi({
        intent,
        tone: els.toneSelect?.value || "standard",
        context: combinedContext,
        instruction:
          intent === "new_email"
            ? "Lag en komplett norsk e-post basert på prompten."
            : intent === "complaint"
              ? "Lag en leverandørklar reklamasjonsmail på norsk med relevant informasjon og tydelig neste steg."
              : "Foreslå smart handlinger for e-posten: flagg, pin, foreslå mappe og kort begrunnelse.",
      });

      const text = extractText(response);
      state.lastTools = text;
      els.toolsOutput.value = text;

      const doneText =
        intent === "new_email"
          ? "Ny e-post klar"
          : intent === "complaint"
            ? "Reklamasjonsutkast klart"
            : "Analyse klar";

      setStatus(doneText, "loading", 1800);
    } catch (error) {
      setStatus("Kunne ikke fullføre verktøyet", "error", 2600);
    } finally {
      setLoading(false);
    }
  }

  async function callAi(payload) {
    const response = await fetch("/api/ai/reply", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      throw new Error(`API feilet med status ${response.status}`);
    }

    return response.json();
  }

  function extractText(data) {
    if (!data) return "";
    if (typeof data === "string") return data.trim();
    if (typeof data.reply === "string") return data.reply.trim();
    if (typeof data.output === "string") return data.output.trim();
    if (typeof data.text === "string") return data.text.trim();
    if (typeof data.content === "string") return data.content.trim();
    return "";
  }

  async function insertIntoMessage(text) {
    try {
      if (!text.trim()) {
        setStatus("Ingen tekst å sette inn", "error", 1800);
        return;
      }

      const item = Office.context?.mailbox?.item;
      if (!item?.body?.setSelectedDataAsync) {
        setStatus("Outlook støtter ikke innsetting her", "error", 2200);
        return;
      }

      item.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setStatus("Svar satt inn i e-post", "loading", 1800);
          return;
        }
        setStatus("Kunne ikke sette inn svar", "error", 2400);
      });
    } catch (error) {
      setStatus("Kunne ikke sette inn svar", "error", 2400);
    }
  }

  async function copyText(text) {
    try {
      if (!text.trim()) {
        setStatus("Ingen tekst å kopiere", "error", 1800);
        return;
      }

      await navigator.clipboard.writeText(text);
      setStatus("Kopiert", "loading", 1400);
    } catch (error) {
      setStatus("Kunne ikke kopiere", "error", 2200);
    }
  }

  function clearOutputs() {
    if (els.replyOutput) els.replyOutput.value = "";
    if (els.summaryOutput) els.summaryOutput.value = "";
    if (els.tasksOutput) els.tasksOutput.value = "";
    if (els.toolsOutput) els.toolsOutput.value = "";
    if (els.toolPrompt) els.toolPrompt.value = "";
    state.lastReply = "";
    state.lastSummary = "";
    state.lastTasks = "";
    state.lastTools = "";
    state.selectedAdjustment = "";
  }

  function ensureContextLoaded() {
    if (!state.currentBodyText) {
      const item = Office.context?.mailbox?.item;
      if (item) {
        state.currentItem = item;
      }
    }
  }

  function setLoading(isLoading, text = "") {
    if (isLoading) {
      els.statusLine?.classList.add("show", "loading");
      els.statusLine?.classList.remove("error");
      if (els.statusText) els.statusText.textContent = text || "Jobber...";
      toggleButtons(true);
      return;
    }

    toggleButtons(false);
    els.statusLine?.classList.remove("loading");
  }

  function setStatus(text, type = "loading", timeout = 0) {
    if (!els.statusLine || !els.statusText) return;

    els.statusLine.classList.add("show");
    els.statusLine.classList.remove("loading", "error");
    els.statusLine.classList.add(type);
    els.statusText.textContent = text;

    if (timeout > 0) {
      window.clearTimeout(setStatus._timer);
      setStatus._timer = window.setTimeout(() => {
        els.statusLine.classList.remove("show", "loading", "error");
        els.statusText.textContent = "";
      }, timeout);
    }
  }

  function toggleButtons(disabled) {
    [
      els.generateReplyBtn,
      els.insertReplyBtn,
      els.copyReplyBtn,
      els.regenerateBtn,
      ...els.promptActions,
      els.generateSummaryBtn,
      els.copySummaryBtn,
      els.generateTasksBtn,
      els.copyTasksBtn,
      els.generateNewEmailBtn,
      els.generateComplaintBtn,
      els.suggestToolsBtn,
      els.copyToolsBtn,
      els.refreshContextBtn,
      els.clearAllBtn,
    ]
      .filter(Boolean)
      .forEach((btn) => {
        btn.disabled = disabled;
        btn.style.opacity = disabled ? "0.65" : "1";
        btn.style.pointerEvents = disabled ? "none" : "auto";
      });
  }
})();