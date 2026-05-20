import {
      firebaseEnv,
      loadFirestoreDoc,
      saveFirestoreDoc,
      serverTimestamp,
      observeAuth,
      signInAnon,
      signInEmail,
      signUpEmail,
      linkEmail,
      signOutUser,
    } from "../firebase.js?v19";

    (async () => {
      const DEFAULT_SLOT_COUNT = 4;
      const SKILL_COUNT = 6;
      const DEFAULT_PARTY_COUNT = 4;
      const FIRESTORE_COLLECTION = "partyRecords";
      const SAVE_DELAY_MS = 600;

      const partyList = document.getElementById("party-list");
      const partyTemplate = document.getElementById("party-template");
      const template = document.getElementById("slot-template");
      const statusText = document.getElementById("status-text");
      const addPartyButton = document.getElementById("add-party");
      const toast = document.getElementById("toast");
      const toastMessage = document.getElementById("toast-message");
      const toastUndo = document.getElementById("toast-undo");
      const authTrigger = document.getElementById("auth-trigger");
      const authMenu = document.getElementById("auth-menu");
      const authTitle = document.getElementById("auth-title");
      const authSubtitle = document.getElementById("auth-subtitle");
      const authForm = document.getElementById("auth-form");
      const authEmailInput = document.getElementById("auth-email");
      const authPasswordInput = document.getElementById("auth-password");
      const authEmailLoginButton = document.getElementById("auth-email-login");
      const authEmailSignupButton = document.getElementById("auth-email-signup");
      const authSignoutButton = document.getElementById("auth-signout");


      const defaultEntry = () => ({
        name: "",
        imageUrl: "",
        skills: Array.from({ length: SKILL_COUNT }, () => ""),
      });

      const normalizeText = (text) => String(text ?? "").replace(/\s+/g, " ").trim();

      const normalizeEntry = (entry) => {
        const skills = Array.from({ length: SKILL_COUNT }, (_, skillIndex) =>
          normalizeText(entry && entry.skills && entry.skills[skillIndex])
        );
        return {
          name: normalizeText(entry && entry.name),
          imageUrl: normalizeText(entry && entry.imageUrl),
          skills,
        };
      };

      const buildParty = (entries) => {
        if (Array.isArray(entries) && entries.length > 0) {
          return entries.map((entry) => normalizeEntry(entry));
        }
        return Array.from({ length: DEFAULT_SLOT_COUNT }, () => defaultEntry());
      };

      const buildDefaultStore = () => ({
        activeParty: 1,
        parties: {},
        partyCount: DEFAULT_PARTY_COUNT,
        partyNames: {},
      });

      const normalizeStore = (raw) => {
        if (!raw || typeof raw !== "object") {
          return buildDefaultStore();
        }
        return {
          activeParty: Number.isFinite(raw.activeParty) ? raw.activeParty : 1,
          parties: raw.parties && typeof raw.parties === "object" ? raw.parties : {},
          partyCount: Number.isFinite(raw.partyCount) ? raw.partyCount : DEFAULT_PARTY_COUNT,
          partyNames: raw.partyNames && typeof raw.partyNames === "object" ? raw.partyNames : {},
        };
      };

      const isStoreEmpty = (store) => {
        if (!store) return true;
        const partyNames = store.partyNames || {};
        if (Object.values(partyNames).some((name) => normalizeText(name))) {
          return false;
        }
        const parties = store.parties || {};
        return !Object.values(parties).some((entries) =>
          Array.isArray(entries) &&
          entries.some((entry) => {
            if (!entry) return false;
            if (normalizeText(entry.name)) return true;
            if (normalizeText(entry.imageUrl)) return true;
            if (Array.isArray(entry.skills) && entry.skills.some((skill) => normalizeText(skill))) {
              return true;
            }
            return false;
          })
        );
      };

      const canWriteStore = (store) => !isStoreEmpty(store);

      const readLocalCache = (uid, { isAnonymous = false } = {}) => {
        if (!window.EldersignLegacyStorage) return null;
        return window.EldersignLegacyStorage.readLegacyStore(uid, {
          normalizeStore,
          isAnonymous,
        });
      };

      const loadStore = async (uid, { isAnonymous = false } = {}) => {
        const localStore = readLocalCache(uid, { isAnonymous });
        try {
          const payload = await loadFirestoreDoc([FIRESTORE_COLLECTION, uid]);
          if (payload && payload.store) {
            const normalized = normalizeStore(payload.store);
            if (localStore && isStoreEmpty(normalized) && !isStoreEmpty(localStore)) {
              return { store: localStore, source: "local" };
            }
            return { store: normalized, source: "firestore" };
          }
        } catch (error) {
          setStatus("Firestoreの読み込みに失敗しました。ローカルを使用します。");
        }

        if (localStore) return { store: localStore, source: "local" };

        return { store: buildDefaultStore(), source: "default" };
      };

      let saveTimer = null;

      const saveStore = (store, { immediate = false } = {}) => {
        const targetUid = storeOwnerUid || currentUid;
        if (!canWriteStore(store)) {
          if (saveTimer) {
            clearTimeout(saveTimer);
            saveTimer = null;
          }
          return;
        }

        if (window.EldersignLegacyStorage) {
          window.EldersignLegacyStorage.writeLegacyCache(targetUid, store);
        }

        if (saveTimer) {
          clearTimeout(saveTimer);
          saveTimer = null;
        }

        const runSave = async () => {
          if (!targetUid || !canWriteStore(store)) return;
          try {
            await saveFirestoreDoc([FIRESTORE_COLLECTION, targetUid], {
              store,
              updatedAt: serverTimestamp(),
              env: firebaseEnv,
            });
          } catch (error) {
            setStatus("Firestoreの保存に失敗しました。");
          }
        };

        if (immediate) {
          runSave();
          return;
        }
        saveTimer = setTimeout(runSave, SAVE_DELAY_MS);
      };

      const copyText = async (text) => {
        if (navigator.clipboard && navigator.clipboard.writeText) {
          await navigator.clipboard.writeText(text);
          return;
        }
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.setAttribute("readonly", "");
        textarea.style.position = "fixed";
        textarea.style.top = "-9999px";
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand("copy");
        document.body.removeChild(textarea);
      };

      const setStatus = (message) => {
        if (!statusText) return;
        statusText.textContent = message;
      };

      const setAuthStatus = (user) => {
        if (!authTitle || !authSubtitle) return;
        if (!user || user.isAnonymous) {
          authTitle.textContent = "未ログイン";
          authSubtitle.textContent = "ログインすると他端末と共有できます";
          return;
        }
        authTitle.textContent = "ログイン中";
        authSubtitle.textContent = user.email || "メールログイン中";
      };

      const updateAuthButtons = (user) => {
        if (!authForm || !authSignoutButton) return;
        if (!user || user.isAnonymous) {
          authForm.style.display = "grid";
          authSignoutButton.style.display = "none";
          if (authTrigger) authTrigger.classList.remove("is-logged-in");
          return;
        }
        authForm.style.display = "none";
        authSignoutButton.style.display = "inline-flex";
        if (authTrigger) authTrigger.classList.add("is-logged-in");
      };

      const toastState = {
        timer: null,
        onUndo: null,
      };

      const showToast = (message, onUndo) => {
        if (!toast || !toastMessage || !toastUndo) return;
        toastMessage.textContent = message;
        toast.classList.add("show");
        toastUndo.style.display = onUndo ? "inline-flex" : "none";
        toastState.onUndo = onUndo || null;
        clearTimeout(toastState.timer);
        toastState.timer = setTimeout(() => {
          toast.classList.remove("show");
          toastState.onUndo = null;
        }, 5000);
      };

      const dragState = {
        type: null,
        slotIndex: null,
        skillIndex: null,
        partyId: null,
      };

      let store = buildDefaultStore();
      let storeOwnerUid = null;
      let currentUid = null;
      let currentUser = null;
      let pendingInit = null;
      let hasAppliedIncoming = false;
      let anonSignInTimer = null;
      const ensureParty = (partyId) => {
        if (!store.parties[partyId]) {
          store.parties[partyId] = buildParty();
        }
        return store.parties[partyId];
      };
      const ensurePartyCount = (count) => {
        const targetCount = Math.max(1, Math.floor(count));
        store.partyCount = targetCount;
        for (let i = 1; i <= targetCount; i += 1) {
          ensureParty(i);
        }
      };
      const clampParty = (partyId) =>
        Math.min(store.partyCount, Math.max(1, Number.parseInt(partyId, 10) || 1));
      const slotRefs = {};

      const updatePreview = (ref, entry) => {
        if (!ref || !ref.preview) return;
        if (entry.imageUrl) {
          ref.preview.classList.add("has-image");
          ref.previewImage.src = entry.imageUrl;
          ref.previewImage.alt = entry.name ? `${entry.name}の画像` : "モンスター画像";
        } else {
          ref.preview.classList.remove("has-image");
          ref.previewImage.removeAttribute("src");
          ref.previewImage.alt = "";
        }
      };

      const getPartyEntries = (partyId) => ensureParty(partyId);

      const renderSlot = (ref, entry, index) => {
        if (!ref || !entry) return;
        ref.nameInput.value = entry.name;
        ref.imageInput.value = entry.imageUrl;
        ref.skillInputs.forEach((input, skillIndex) => {
          input.value = entry.skills[skillIndex] || "";
        });
        updatePreview(ref, entry);
        ref.card.classList.remove("is-editing");
        ref.card.dataset.index = String(index);
      };

      const isEntryEmpty = (entry) =>
        !entry.name && !entry.imageUrl && entry.skills.every((skill) => !skill);

      const resolveSlot = (partyId, requestedSlot) => {
        const entries = getPartyEntries(partyId);
        if (Number.isFinite(requestedSlot) && requestedSlot >= 0) {
          while (entries.length <= requestedSlot) {
            entries.push(defaultEntry());
          }
          return requestedSlot;
        }
        const emptyIndex = entries.findIndex((entry) => isEntryEmpty(entry));
        return emptyIndex >= 0 ? emptyIndex : entries.length;
      };

      const readIncoming = () => {
        const params = new URLSearchParams(window.location.search);
        if (![...params.keys()].length) return null;

        const name = normalizeText(params.get("name"));
        const imageUrl = normalizeText(params.get("img"));
        const skillParam = params.get("skills");
        const skills = Array.from({ length: SKILL_COUNT }, () => "");
        const assignSkillRange = (values, startIndex) => {
          values.forEach((skill, index) => {
            const targetIndex = startIndex + index;
            if (targetIndex >= SKILL_COUNT) return;
            const normalized = normalizeText(skill);
            if (normalized) skills[targetIndex] = normalized;
          });
        };

        if (skillParam) {
          assignSkillRange(skillParam.split("|"), 1);
        }

        const skillList = params.getAll("skill");
        if (skillList.length) {
          assignSkillRange(skillList, 1);
        }

        for (let i = 0; i < SKILL_COUNT; i += 1) {
          const skillValue = normalizeText(params.get(`skill${i}`));
          if (skillValue) skills[i] = skillValue;
        }

        if (!name && !imageUrl && !skills.some((skill) => skill)) return null;

        const slotValue = Number.parseInt(params.get("slot"), 10);
        const requestedSlot = Number.isFinite(slotValue) ? slotValue - 1 : null;

        const partyValue = Number.parseInt(params.get("party"), 10);
        const requestedParty = Number.isFinite(partyValue) ? Math.max(1, Math.floor(partyValue)) : null;

        return { name, imageUrl, skills, requestedSlot, requestedParty };
      };

      const applyIncoming = (incoming) => {
        if (!incoming) return;
        const targetParty = incoming.requestedParty || 1;
        if (targetParty > store.partyCount) {
          ensurePartyCount(targetParty);
        }
        const slotIndex = resolveSlot(targetParty, incoming.requestedSlot);
        if (slotIndex === null) {
          setStatus("空き枠がありません。手動で入れ替えてください。");
          return;
        }
        const entry = defaultEntry();
        entry.name = incoming.name;
        entry.imageUrl = incoming.imageUrl;
        entry.skills = Array.from({ length: SKILL_COUNT }, (_, index) => incoming.skills[index] || "");
        const entries = getPartyEntries(targetParty);
        entries[slotIndex] = entry;
        saveStore(store);
        renderAll();
        setStatus(`PT${targetParty}に追加しました`);
        if (history.replaceState) {
          history.replaceState(null, "", window.location.pathname);
        }
      };

      const handleInput = (event) => {
        const input = event.target;
        const slotIndex = Number.parseInt(input.dataset.slot, 10);
        const partyId = clampParty(input.dataset.party);
        const entries = getPartyEntries(partyId);
        if (!Number.isFinite(slotIndex) || !entries[slotIndex]) return;

        const entry = entries[slotIndex];
        const value = normalizeText(input.value);
        if (input.dataset.field === "name") {
          entry.name = value;
          const ref = slotRefs[partyId] ? slotRefs[partyId][slotIndex] : null;
          if (ref) updatePreview(ref, entry);
        } else if (input.dataset.field === "imageUrl") {
          entry.imageUrl = value;
          const ref = slotRefs[partyId] ? slotRefs[partyId][slotIndex] : null;
          if (ref) updatePreview(ref, entry);
        } else if (input.dataset.field === "skill") {
          const skillIndex = Number.parseInt(input.dataset.skillIndex, 10);
          if (Number.isFinite(skillIndex)) {
            entry.skills[skillIndex] = value;
          }
        }

        saveStore(store);
      };

      const duplicateSlot = (partyId, index) => {
        const entries = getPartyEntries(partyId);
        const target = entries[index];
        if (!target) return;
        const cloned = {
          name: target.name,
          imageUrl: target.imageUrl,
          skills: Array.from({ length: SKILL_COUNT }, (_, skillIndex) => target.skills[skillIndex] || ""),
        };
        entries.splice(index + 1, 0, cloned);
        saveStore(store);
        renderAll();
        setStatus(`PT${partyId}の枠を複製しました`);
      };

      const clearSlot = (partyId, index) => {
        const entries = getPartyEntries(partyId);
        const removed = entries.splice(index, 1)[0];
        if (entries.length === 0) {
          entries.push(defaultEntry());
        }
        saveStore(store);
        renderAll();
        setStatus(`PT${partyId}の枠を削除しました`);
        if (removed) {
          showToast("枠を削除しました", () => {
            const restoreEntries = getPartyEntries(partyId);
            const insertIndex = Math.min(index, restoreEntries.length);
            restoreEntries.splice(insertIndex, 0, removed);
            saveStore(store);
            renderAll();
            setStatus("削除を取り消しました");
          });
        }
      };

      const clearAll = (partyId) => {
        const confirmed = window.confirm(`PT${partyId}を削除します。よろしいですか？`);
        if (!confirmed) return;
        const targetId = Math.max(1, Math.floor(partyId));
        if (store.partyCount <= 1) {
          const entries = getPartyEntries(1);
          entries.length = 0;
          for (let i = 0; i < DEFAULT_SLOT_COUNT; i += 1) {
            entries.push(defaultEntry());
          }
          saveStore(store);
          setStatus("PT1のみなので枠を初期化しました");
          renderAll();
          return;
        }
        delete store.parties[targetId];
        delete store.partyNames[String(targetId)];
        for (let id = targetId + 1; id <= store.partyCount; id += 1) {
          if (store.parties[id]) {
            store.parties[id - 1] = store.parties[id];
            delete store.parties[id];
          }
          const name = store.partyNames[String(id)];
          if (name != null) {
            store.partyNames[String(id - 1)] = name;
            delete store.partyNames[String(id)];
          }
        }
        store.partyCount = Math.max(1, store.partyCount - 1);
        store.activeParty = clampParty(store.activeParty);
        saveStore(store);
        setStatus(`PT${targetId}を削除しました`);
        renderAll();
      };

      const reorderParties = (fromPartyId, toPartyId) => {
        const fromId = clampParty(fromPartyId);
        const toId = clampParty(toPartyId);
        if (fromId === toId) return false;

        const ordered = [];
        for (let id = 1; id <= store.partyCount; id += 1) {
          ordered.push({
            id,
            entries: getPartyEntries(id).map((entry) => normalizeEntry(entry)),
            name: store.partyNames[String(id)] || "",
          });
        }

        const fromIndex = fromId - 1;
        const toIndex = toId - 1;
        const [moved] = ordered.splice(fromIndex, 1);
        if (!moved) return false;
        ordered.splice(toIndex, 0, moved);

        store.parties = {};
        store.partyNames = {};
        ordered.forEach((party, index) => {
          const nextId = index + 1;
          store.parties[nextId] = party.entries;
          if (party.name) {
            store.partyNames[String(nextId)] = party.name;
          }
        });
        store.partyCount = ordered.length;
        return true;
      };

      const addSlot = (partyId) => {
        const entries = getPartyEntries(partyId);
        entries.push(defaultEntry());
        saveStore(store);
        renderAll();
        setStatus(`PT${partyId}に枠を追加しました`);
      };

      const addParty = () => {
        const nextPartyId = store.partyCount + 1;
        ensurePartyCount(nextPartyId);
        saveStore(store);
        renderAll();
        setStatus(`PT${nextPartyId}を追加しました`);
      };

      const openImageEditor = (partyId, index) => {
        const ref = slotRefs[partyId] ? slotRefs[partyId][index] : null;
        if (!ref) return;
        ref.card.classList.add("is-editing");
        ref.imageInput.focus();
        ref.imageInput.select();
      };

      const closeImageEditor = (partyId, index) => {
        const ref = slotRefs[partyId] ? slotRefs[partyId][index] : null;
        if (!ref) return;
        ref.card.classList.remove("is-editing");
      };

      const swapSlots = (fromPartyId, fromIndex, toPartyId, toIndex) => {
        if (fromPartyId === toPartyId && fromIndex === toIndex) return;
        const fromEntries = getPartyEntries(fromPartyId);
        const toEntries = getPartyEntries(toPartyId);
        if (fromPartyId === toPartyId) {
          const temp = fromEntries[fromIndex];
          fromEntries[fromIndex] = toEntries[toIndex];
          toEntries[toIndex] = temp;
          saveStore(store);
          renderAll();
          setStatus(`PT${fromPartyId}の枠を入れ替えました`);
          return;
        }

        const [moved] = fromEntries.splice(fromIndex, 1);
        if (!moved) return;
        if (fromEntries.length === 0) {
          fromEntries.push(defaultEntry());
        }
        const insertIndex = Math.max(0, Math.min(toIndex, toEntries.length));
        toEntries.splice(insertIndex, 0, moved);
        saveStore(store);
        renderAll();
        setStatus(`PT${fromPartyId}からPT${toPartyId}へ枠を移動しました`);
      };

      const swapSkills = (partyId, slotIndex, fromIndex, toIndex) => {
        if (fromIndex === toIndex) return;
        const entries = getPartyEntries(partyId);
        const entry = entries[slotIndex];
        if (!entry) return;
        const temp = entry.skills[fromIndex];
        entry.skills[fromIndex] = entry.skills[toIndex];
        entry.skills[toIndex] = temp;
        saveStore(store);
        const ref = slotRefs[partyId] ? slotRefs[partyId][slotIndex] : null;
        if (ref) renderSlot(ref, entry, slotIndex);
        setStatus(`PT${partyId}の枠${slotIndex + 1}でスキルを入れ替えました`);
      };

      const clearDragTarget = () => {
        document.querySelectorAll(".drag-target").forEach((el) => {
          el.classList.remove("drag-target");
        });
      };

      const onDragEnd = () => {
        dragState.type = null;
        dragState.slotIndex = null;
        dragState.skillIndex = null;
        dragState.partyId = null;
        clearDragTarget();
      };

      const blockDrag = (event) => {
        event.preventDefault();
        event.stopPropagation();
      };

      const buildSlot = (entry, partyId, index, grid) => {
        const clone = template.content.cloneNode(true);
        const card = clone.querySelector(".slot-card");
        const duplicateButton = clone.querySelector(".slot-duplicate");
        const clearButton = clone.querySelector(".slot-clear:not(.slot-duplicate)");
        const nameInput = clone.querySelector(".slot-name");
        const imageInput = clone.querySelector(".slot-image-url");
        const preview = clone.querySelector(".slot-preview");
        const previewImage = clone.querySelector(".slot-image");
        const skillRows = [...clone.querySelectorAll(".skill-row")];
        const skillInputs = [...clone.querySelectorAll(".slot-skill")];

        nameInput.dataset.party = String(partyId);
        nameInput.dataset.slot = String(index);
        nameInput.dataset.field = "name";
        imageInput.dataset.party = String(partyId);
        imageInput.dataset.slot = String(index);
        imageInput.dataset.field = "imageUrl";
        skillInputs.forEach((input, skillIndex) => {
          input.dataset.party = String(partyId);
          input.dataset.slot = String(index);
          input.dataset.field = "skill";
          input.dataset.skillIndex = String(skillIndex);
        });

        nameInput.addEventListener("input", handleInput);
        nameInput.addEventListener("dragstart", blockDrag);
        imageInput.addEventListener("input", handleInput);
        imageInput.addEventListener("dragstart", blockDrag);
        imageInput.addEventListener("blur", () => closeImageEditor(partyId, index));
        imageInput.addEventListener("keydown", (event) => {
          if (event.key === "Escape") {
            event.preventDefault();
            closeImageEditor(partyId, index);
          }
        });
        skillInputs.forEach((input) => {
          input.addEventListener("input", handleInput);
          input.addEventListener("dragstart", blockDrag);
        });
        if (duplicateButton) {
          duplicateButton.addEventListener("click", () => duplicateSlot(partyId, index));
        }
        clearButton.addEventListener("click", () => clearSlot(partyId, index));
        preview.addEventListener("click", () => openImageEditor(partyId, index));
        preview.addEventListener("keydown", (event) => {
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            openImageEditor(partyId, index);
          }
        });

        card.setAttribute("draggable", "false");
        const allowCardDrag = (event) => {
          if (event.target.closest("input, textarea, select, button, label, .field")) {
            card.setAttribute("draggable", "false");
            return;
          }
          card.setAttribute("draggable", "true");
        };
        const disableCardDrag = () => {
          card.setAttribute("draggable", "false");
        };
        card.addEventListener("mousedown", allowCardDrag);
        card.addEventListener("mouseup", disableCardDrag);
        card.addEventListener("mouseleave", disableCardDrag);
        card.addEventListener("dragstart", (event) => {
          if (event.target.closest("input, textarea, select, button, label, .field")) {
            event.preventDefault();
            return;
          }
          dragState.type = "slot";
          dragState.partyId = partyId;
          dragState.slotIndex = index;
          dragState.skillIndex = null;
          card.classList.add("is-dragging");
          event.dataTransfer.effectAllowed = "move";
          event.dataTransfer.setData("text/plain", "slot");
        });
        card.addEventListener("dragend", () => {
          card.classList.remove("is-dragging");
          onDragEnd();
          disableCardDrag();
        });
        card.addEventListener("dragover", (event) => {
          if (dragState.type !== "slot") return;
          event.preventDefault();
          card.classList.add("drag-target");
        });
        card.addEventListener("dragleave", () => {
          card.classList.remove("drag-target");
        });
        card.addEventListener("drop", (event) => {
          if (dragState.type !== "slot") return;
          event.preventDefault();
          card.classList.remove("drag-target");
          if (dragState.slotIndex === null) return;
          if (!dragState.partyId) return;
          swapSlots(dragState.partyId, dragState.slotIndex, partyId, index);
          onDragEnd();
        });

        skillRows.forEach((row, skillIndex) => {
          row.dataset.party = String(partyId);
          row.dataset.slot = String(index);
          row.dataset.skillIndex = String(skillIndex);
          const handle = row.querySelector(".skill-drag");
          if (!handle) {
            return;
          }
          handle.dataset.party = String(partyId);
          handle.dataset.slot = String(index);
          handle.dataset.skillIndex = String(skillIndex);
          handle.addEventListener("dragstart", (event) => {
            event.stopPropagation();
            disableCardDrag();
            dragState.type = "skill";
            dragState.partyId = partyId;
            dragState.slotIndex = index;
            dragState.skillIndex = skillIndex;
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", "skill");
          });
          handle.addEventListener("dragend", onDragEnd);
          row.addEventListener("dragover", (event) => {
            if (dragState.type !== "skill") return;
            if (dragState.partyId !== partyId || dragState.slotIndex !== index) return;
            event.preventDefault();
            row.classList.add("drag-target");
          });
          row.addEventListener("dragleave", () => {
            row.classList.remove("drag-target");
          });
          row.addEventListener("drop", (event) => {
            if (dragState.type !== "skill") return;
            if (dragState.partyId !== partyId || dragState.slotIndex !== index) return;
            event.preventDefault();
            row.classList.remove("drag-target");
            if (dragState.skillIndex === null) return;
            swapSkills(partyId, index, dragState.skillIndex, skillIndex);
            onDragEnd();
          });
        });

        const ref = {
          card,
          nameInput,
          imageInput,
          preview,
          previewImage,
          skillInputs,
        };
        if (!slotRefs[partyId]) {
          slotRefs[partyId] = [];
        }
        slotRefs[partyId].push(ref);
        renderSlot(ref, entry, index);
        grid.appendChild(clone);
      };

      const renderAll = () => {
        partyList.innerHTML = "";
        Object.keys(slotRefs).forEach((key) => delete slotRefs[key]);
        for (let partyId = 1; partyId <= store.partyCount; partyId += 1) {
          const entries = getPartyEntries(partyId);
          const clone = partyTemplate.content.cloneNode(true);
          const section = clone.querySelector(".party-section");
          const header = clone.querySelector(".party-header");
          const nameInput = clone.querySelector(".party-name");
          const grid = clone.querySelector(".party-grid");
          const toggleButton = clone.querySelector('[data-action="toggle"]');
          const actions = clone.querySelectorAll("[data-action]");
          const defaultPartyName = `PT ${partyId}`;
          if (section) {
            section.dataset.party = String(partyId);
            section.classList.remove("is-collapsed");
            section.addEventListener("dragover", (event) => {
              if (dragState.type !== "party") return;
              event.preventDefault();
              section.classList.add("drag-target");
            });
            section.addEventListener("dragleave", () => {
              section.classList.remove("drag-target");
            });
            section.addEventListener("drop", (event) => {
              if (dragState.type !== "party") return;
              event.preventDefault();
              section.classList.remove("drag-target");
              if (!dragState.partyId) return;
              if (!reorderParties(dragState.partyId, partyId)) {
                onDragEnd();
                return;
              }
              saveStore(store);
              renderAll();
              setStatus(`PT${dragState.partyId}をPT${partyId}の位置へ移動しました`);
              onDragEnd();
            });
          }
          if (header) {
            header.setAttribute("aria-expanded", "true");
            const toggleSection = () => {
              if (!section) return;
              const nextCollapsed = !section.classList.contains("is-collapsed");
              section.classList.toggle("is-collapsed", nextCollapsed);
              header.setAttribute("aria-expanded", nextCollapsed ? "false" : "true");
            };
            header.setAttribute("draggable", "true");
            header.addEventListener("dragstart", (event) => {
              if (event.target.closest("input, button")) {
                event.preventDefault();
                return;
              }
              dragState.type = "party";
              dragState.partyId = partyId;
              dragState.slotIndex = null;
              dragState.skillIndex = null;
              if (section) {
                section.classList.add("is-dragging");
              }
              event.dataTransfer.effectAllowed = "move";
              event.dataTransfer.setData("text/plain", "party");
            });
            header.addEventListener("dragend", () => {
              if (section) {
                section.classList.remove("is-dragging");
              }
              onDragEnd();
            });
            header.addEventListener("click", (event) => {
              if (event.target.closest("input")) return;
              toggleSection();
            });
            header.addEventListener("keydown", (event) => {
              if (event.key !== "Enter" && event.key !== " ") return;
              if (event.target.closest("input")) return;
              event.preventDefault();
              toggleSection();
            });
            if (toggleButton) {
              toggleButton.addEventListener("click", (event) => {
                event.stopPropagation();
                toggleSection();
              });
            }
          }
          if (nameInput) {
            nameInput.placeholder = defaultPartyName;
            nameInput.value = store.partyNames[String(partyId)] || "";
            nameInput.addEventListener("click", (event) => event.stopPropagation());
            nameInput.addEventListener("keydown", (event) => event.stopPropagation());
            nameInput.addEventListener("input", () => {
              store.partyNames[String(partyId)] = nameInput.value.trim();
              saveStore(store);
            });
          }
          actions.forEach((button) => {
            button.dataset.party = String(partyId);
            button.addEventListener("click", (event) => {
              event.stopPropagation();
            });
            const action = button.dataset.action;
            if (action === "add") {
              button.addEventListener("click", () => {
                addSlot(partyId);
              });
            }
            if (action === "clear") {
              button.addEventListener("click", () => {
                clearAll(partyId);
              });
            }
          });

          entries.forEach((entry, index) => {
            buildSlot(entry, partyId, index, grid);
          });
          partyList.appendChild(clone);
        }
      };

      const initStore = async (uid, { isAnonymous = false } = {}) => {
        if (!uid) return;
        setStatus("Firestoreから読み込み中...");
        const { store: loadedStore, source } = await loadStore(uid, { isAnonymous });
        store = loadedStore;
        storeOwnerUid = uid;
        if (!Number.isFinite(store.partyCount) || store.partyCount < 1) {
          store.partyCount = DEFAULT_PARTY_COUNT;
        }
        if (!store.partyNames || typeof store.partyNames !== "object") {
          store.partyNames = {};
        }
        ensurePartyCount(store.partyCount);
        store.activeParty = clampParty(store.activeParty);
        renderAll();
        if (!hasAppliedIncoming) {
          applyIncoming(readIncoming());
          hasAppliedIncoming = true;
        }
        if (source !== "firestore") {
          saveStore(store, { immediate: true });
        }
        setStatus("保存は自動です");
      };

      const initLocalStore = () => {
        store = readLocalCache(null, { isAnonymous: true }) || buildDefaultStore();
        storeOwnerUid = null;
        if (!Number.isFinite(store.partyCount) || store.partyCount < 1) {
          store.partyCount = DEFAULT_PARTY_COUNT;
        }
        if (!store.partyNames || typeof store.partyNames !== "object") {
          store.partyNames = {};
        }
        ensurePartyCount(store.partyCount);
        store.activeParty = clampParty(store.activeParty);
        renderAll();
        if (!hasAppliedIncoming) {
          applyIncoming(readIncoming());
          hasAppliedIncoming = true;
        }
        setStatus("ログインすると他端末と共有できます");
      };

      const migrateIfNeeded = async (fromUid, toUid) => {
        if (!fromUid || !toUid || fromUid === toUid) return;
        const fromCache = readLocalCache(fromUid, { isAnonymous: true });
        if (!fromCache || !canWriteStore(fromCache)) return;
        const remote = await loadStore(toUid, { isAnonymous: false });
        if (remote && remote.store && remote.source === "firestore") return;
        try {
          await saveFirestoreDoc([FIRESTORE_COLLECTION, toUid], {
            store: fromCache,
            updatedAt: serverTimestamp(),
            env: firebaseEnv,
          });
        } catch (error) {
          setStatus("ログイン後の引き継ぎに失敗しました。");
        }
      };

      const handleAuthChange = async (user) => {
        if (anonSignInTimer) {
          clearTimeout(anonSignInTimer);
          anonSignInTimer = null;
        }
        currentUser = user;
        currentUid = user ? user.uid : null;
        setAuthStatus(user);
        updateAuthButtons(user);
        if (!currentUid) return;
        if (pendingInit) {
          pendingInit.cancelled = true;
        }
        const ticket = { cancelled: false };
        pendingInit = ticket;
        await initStore(currentUid, { isAnonymous: user.isAnonymous });
        if (ticket.cancelled) return;
      };

      if (toastUndo) {
        toastUndo.addEventListener("click", () => {
          if (!toastState.onUndo) return;
          const handler = toastState.onUndo;
          toastState.onUndo = null;
          toast.classList.remove("show");
          handler();
        });
      }
      if (addPartyButton) {
        addPartyButton.addEventListener("click", addParty);
      }
      const readAuthInputs = () => {
        if (!authEmailInput || !authPasswordInput) return null;
        const email = authEmailInput.value.trim();
        const password = authPasswordInput.value;
        if (!email) {
          setStatus("メールアドレスを入力してください。");
          return null;
        }
        return { email, password };
      };

      if (authEmailLoginButton) {
        authEmailLoginButton.addEventListener("click", async () => {
          const input = readAuthInputs();
          if (!input) return;
          if (!input.password) {
            setStatus("パスワードを入力してください。");
            return;
          }
          try {
            let result;
            if (currentUser && currentUser.isAnonymous) {
              result = await linkEmail(input.email, input.password);
            } else {
              result = await signInEmail(input.email, input.password);
            }
            if (result && result.user && currentUid && result.user.uid !== currentUid) {
              await migrateIfNeeded(currentUid, result.user.uid);
              await initStore(result.user.uid, { isAnonymous: false });
            }
            setStatus("ログインしました。");
            authPasswordInput.value = "";
          } catch (error) {
            const code = error && error.code;
            if (currentUser && currentUser.isAnonymous && code === "auth/email-already-in-use") {
              try {
                const result = await signInEmail(input.email, input.password);
                if (result && result.user && currentUid && result.user.uid !== currentUid) {
                  await migrateIfNeeded(currentUid, result.user.uid);
                  await initStore(result.user.uid, { isAnonymous: false });
                }
                setStatus("ログインしました。");
                authPasswordInput.value = "";
                return;
              } catch (signinError) {
                const detail = signinError && (signinError.code || signinError.message);
                console.error("メールログイン失敗", signinError);
                setStatus(`メールログインに失敗しました。${detail || ""}`);
                return;
              }
            }
            const detail = error && (error.code || error.message);
            console.error("メールログイン失敗", error);
            setStatus(`メールログインに失敗しました。${detail || ""}`);
          }
        });
      }

      if (authEmailSignupButton) {
        authEmailSignupButton.addEventListener("click", async () => {
          const input = readAuthInputs();
          if (!input) return;
          if (!input.password) {
            setStatus("パスワードを入力してください。");
            return;
          }
          try {
            let result;
            if (currentUser && currentUser.isAnonymous) {
              result = await linkEmail(input.email, input.password);
            } else {
              result = await signUpEmail(input.email, input.password);
            }
            if (result && result.user && currentUid && result.user.uid !== currentUid) {
              await migrateIfNeeded(currentUid, result.user.uid);
              await initStore(result.user.uid, { isAnonymous: false });
            }
            if (result && result.user) {
              await handleAuthChange(result.user);
            }
            setStatus("アカウントを作成しました。");
            authPasswordInput.value = "";
          } catch (error) {
            const detail = error && (error.code || error.message);
            console.error("新規登録失敗", error);
            setStatus(`新規登録に失敗しました。${detail || ""}`);
          }
        });
      }

      const closeAuthMenu = () => {
        if (!authMenu || !authTrigger) return;
        authMenu.classList.remove("is-open");
        authTrigger.setAttribute("aria-expanded", "false");
      };

      const openAuthMenu = () => {
        if (!authMenu || !authTrigger) return;
        authMenu.classList.add("is-open");
        authTrigger.setAttribute("aria-expanded", "true");
      };

      if (authTrigger) {
        authTrigger.addEventListener("click", (event) => {
          event.stopPropagation();
          if (!authMenu) return;
          if (authMenu.classList.contains("is-open")) {
            closeAuthMenu();
          } else {
            openAuthMenu();
          }
        });
      }

      document.addEventListener("click", (event) => {
        if (!authMenu || !authMenu.classList.contains("is-open")) return;
        const target = event.target;
        if (authMenu.contains(target) || (authTrigger && authTrigger.contains(target))) return;
        closeAuthMenu();
      });

      document.addEventListener("keydown", (event) => {
        if (event.key !== "Escape") return;
        if (!authMenu || !authMenu.classList.contains("is-open")) return;
        closeAuthMenu();
      });
      if (authSignoutButton) {
        authSignoutButton.addEventListener("click", async () => {
          try {
            if (currentUser && currentUser.isAnonymous) {
              setStatus("匿名利用中はログアウトできません");
              return;
            }
            await signOutUser();
          } catch (error) {
            setStatus("ログアウトに失敗しました。");
          }
        });
      }

      observeAuth(async (user) => {
        if (!user) {
          currentUser = null;
          currentUid = null;
          storeOwnerUid = null;
          if (anonSignInTimer) {
            clearTimeout(anonSignInTimer);
          }
          anonSignInTimer = setTimeout(async () => {
            try {
              setStatus("匿名ログイン中...");
              await signInAnon();
            } catch (error) {
              console.error("匿名ログイン失敗", error);
              const detail = error && (error.code || error.message);
              setStatus(`匿名ログインに失敗しました。${detail || ""}`);
            }
          }, 300);
          return;
        }
        handleAuthChange(user);
      });
    })();
