(function () {
  "use strict";

  const API_BASE = "";

  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("file-input");
  const dropzoneContent = document.getElementById("dropzone-content");
  const btnUpload = document.getElementById("btn-upload");
  const btnClear = document.getElementById("btn-clear");
  const uploadProgress = document.getElementById("upload-progress");
  const progressFill = document.getElementById("progress-fill");
  const progressText = document.getElementById("progress-text");
  const messageEl = document.getElementById("message");
  const fileItems = document.getElementById("file-items");
  const emptyState = document.getElementById("empty-state");
  const btnRefresh = document.getElementById("btn-refresh");
  const btnLoadEnregistrements = document.getElementById("btn-load-enregistrements");
  const enregistrementsUploads = document.getElementById("enregistrements-uploads");
  const enregistrementsExports = document.getElementById("enregistrements-exports");
  const uploadsReadingEmpty = document.getElementById("uploads-reading-empty");
  const exportsReadingEmpty = document.getElementById("exports-reading-empty");
  const splitSelect = document.getElementById("split-select");
  const btnSplit = document.getElementById("btn-split");
  const splitProgress = document.getElementById("split-progress");
  const splitResult = document.getElementById("split-result");
  const splitResultTitle = document.getElementById("split-result-title");
  const splitDownloads = document.getElementById("split-downloads");
  const splitMessage = document.getElementById("split-message");
  const btnMergeSplit = document.getElementById("btn-merge-split");
  var currentSplitSourceStem = "";
  var currentSplitDevisList = [];

  function updateMergeButtonState() {
    if (!btnMergeSplit) return;
    var checked = splitDownloads && splitDownloads.querySelectorAll(".split-merge-cb:checked");
    var n = checked ? checked.length : 0;
    btnMergeSplit.disabled = n < 2;
    btnMergeSplit.title = n < 2 ? "Sélectionnez au moins 2 PDF à fusionner" : "Fusionner les " + n + " PDF sélectionnés en un seul fichier (les originaux seront supprimés)";
  }

  function renderSplitDownloads(sourceStem, devisList) {
    currentSplitSourceStem = sourceStem || currentSplitSourceStem;
    currentSplitDevisList = devisList || [];
    var stem = currentSplitSourceStem;
    splitDownloads.innerHTML = (currentSplitDevisList || []).map(function (d, i) {
      var label = d.devis_number != null ? d.devis_number : (d.filename || "").replace(/\.pdf$/i, "");
      var url = API_BASE + (d.download_url || stem + "/" + (d.filename || "")).split("/").map(encodeURIComponent).join("/");
      var viewUrl = url + (url.indexOf("?") >= 0 ? "&" : "?") + "display=1";
      var disabled = i === (currentSplitDevisList.length - 1) ? " disabled title='Aucun PDF suivant'" : " title='Prendre la dernière page du PDF suivant et la mettre en première page de ce devis'";
      var retraiterBtn = "<button type='button' class='btn btn-retraiter'" + disabled + " data-filename='" + escapeHtml(d.filename || "") + "'>Retraiter</button> ";
      var supprimerBtn = "<button type='button' class='btn btn-supprimer btn-ghost' title='Supprimer ce devis' data-filename='" + escapeHtml(d.filename || "") + "'>Supprimer</button> ";
      var cb = "<input type='checkbox' class='split-merge-cb' data-filename='" + escapeHtml(d.filename || "") + "' title='Cocher pour fusionner ce PDF' /> ";
      return (
        "<li>" + cb + retraiterBtn + supprimerBtn +
        "<a href='" + viewUrl + "' target='_blank' rel='noopener' title='Afficher le PDF'>" + escapeHtml(label) + "</a> · " +
        "<span class='size'>(" + (d.page_count != null ? d.page_count : "?") + " page(s))</span> · " +
        "<a href='" + url + "' download title='Télécharger sur le disque'>Télécharger</a></li>"
      );
    }).join("");
    updateMergeButtonState();
    splitDownloads.querySelectorAll(".split-merge-cb").forEach(function (cb) {
      cb.addEventListener("change", updateMergeButtonState);
    });
  }

  let selectedFiles = [];
  let uploadedFilesList = [];

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + " o";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " Ko";
    return (bytes / (1024 * 1024)).toFixed(1) + " Mo";
  }

  function showMessage(text, type) {
    messageEl.textContent = text;
    messageEl.className = "message " + (type === "error" ? "error" : "success");
    messageEl.classList.remove("hidden");
  }

  function hideMessage() {
    messageEl.classList.add("hidden");
  }

  function setProgress(percent, text) {
    uploadProgress.classList.remove("hidden");
    progressFill.style.width = percent + "%";
    progressText.textContent = text || "Envoi en cours…";
  }

  function hideProgress() {
    uploadProgress.classList.add("hidden");
    progressFill.style.width = "0%";
  }

  function updateSelectionUI() {
    if (selectedFiles.length === 0) {
      btnUpload.disabled = true;
      let prev = dropzone.querySelector(".selected-files");
      if (prev) prev.remove();
      return;
    }
    btnUpload.disabled = false;
    let block = dropzone.querySelector(".selected-files");
    if (!block) {
      block = document.createElement("div");
      block.className = "selected-files";
      dropzoneContent.after(block);
    }
    block.innerHTML =
      "<ul>" +
      selectedFiles
        .map(
          (f) =>
            "<li>" +
            escapeHtml(f.name) +
            " <span class='size'>(" +
            formatSize(f.size) +
            ")</span></li>"
        )
        .join("") +
      "</ul>";
  }

  function escapeHtml(s) {
    const div = document.createElement("div");
    div.textContent = s;
    return div.innerHTML;
  }

  function clearSelection() {
    selectedFiles = [];
    fileInput.value = "";
    updateSelectionUI();
    hideMessage();
  }

  dropzone.addEventListener("click", function (e) {
    if (e.target !== fileInput) fileInput.click();
  });

  dropzone.addEventListener("dragover", function (e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add("drag-over");
  });

  dropzone.addEventListener("dragleave", function (e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove("drag-over");
  });

  dropzone.addEventListener("drop", function (e) {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove("drag-over");
    const files = Array.from(e.dataTransfer.files).filter(function (f) {
      return f.name.toLowerCase().endsWith(".pdf");
    });
    if (files.length) {
      selectedFiles = selectedFiles.concat(files);
      updateSelectionUI();
      hideMessage();
    }
  });

  fileInput.addEventListener("change", function () {
    const files = Array.from(fileInput.files || []);
    selectedFiles = selectedFiles.concat(files);
    updateSelectionUI();
    hideMessage();
  });

  btnClear.addEventListener("click", clearSelection);

  btnUpload.addEventListener("click", async function doUpload(overwrite) {
    if (selectedFiles.length === 0) return;
    hideMessage();
    setProgress(0, "Envoi de " + selectedFiles.length + " fichier(s)…");

    const formData = new FormData();
    selectedFiles.forEach(function (f) {
      formData.append("files", f);
    });

    var url = API_BASE + "/api/upload-multiple";
    if (overwrite) url += "?overwrite=true";

    try {
      const res = await fetch(url, {
        method: "POST",
        body: formData,
      });
      var data = await res.json().catch(function () {
        return { detail: "Erreur inconnue" };
      });

      if (res.status === 409) {
        setProgress(0, "");
        hideProgress();
        var detail = data.detail || {};
        var msg = typeof detail === "string" ? detail : (detail.message || "Fichier(s) déjà existant(s).");
        var existingList = (detail.existing || []).join(", ");
        var confirmMsg = "Un ou plusieurs fichiers existent déjà :\n" + existingList + "\n\nVoulez-vous les écraser ?";
        if (window.confirm(confirmMsg)) {
          doUpload(true);
        } else {
          showMessage("Téléversement annulé. Les fichiers existants n'ont pas été modifiés.", "error");
        }
        return;
      }

      setProgress(100, "Envoi terminé.");
      const ok = (data.results || []).filter(function (r) {
        return r.success;
      }).length;
      const err = (data.results || []).filter(function (r) {
        return !r.success && r.error;
      });
      if (err.length) {
        showMessage(
          ok +
            " fichier(s) envoyé(s). Erreurs : " +
            err.map(function (e) {
              return e.filename + " — " + e.error;
            }).join(" ; "),
          "error"
        );
      } else {
        showMessage(ok + " fichier(s) téléversé(s) avec succès.", "success");
      }
      clearSelection();
      loadUploadedList();
    } catch (e) {
      setProgress(0, "");
      hideProgress();
      showMessage("Erreur réseau : " + e.message, "error");
    }

    setTimeout(hideProgress, 1500);
  });

  async function loadUploadedList() {
    try {
      const res = await fetch(API_BASE + "/api/uploads");
      const data = await res.json();
      const files = data.files || [];
      uploadedFilesList = files;
      emptyState.classList.toggle("hidden", files.length > 0);
      fileItems.innerHTML = files
        .map(
          function (f) {
            return (
              "<li><span>" +
              escapeHtml(f.name) +
              "</span> <span class='size'>" +
              formatSize(f.size) +
              "</span></li>"
            );
          }
        )
        .join("");
      var opt = splitSelect.querySelector('option[value=""]');
      splitSelect.innerHTML = "";
      if (opt) splitSelect.appendChild(opt);
      files.forEach(function (f) {
        var o = document.createElement("option");
        o.value = f.name;
        o.textContent = f.name;
        splitSelect.appendChild(o);
      });
      btnSplit.disabled = !splitSelect.value;
    } catch (_) {
      emptyState.textContent = "Impossible de charger la liste.";
      emptyState.classList.remove("hidden");
      fileItems.innerHTML = "";
    }
  }

  splitSelect.addEventListener("change", function () {
    btnSplit.disabled = !splitSelect.value;
  });

  btnSplit.addEventListener("click", async function () {
    var filename = splitSelect.value;
    if (!filename) return;
    splitMessage.classList.add("hidden");
    splitResult.classList.add("hidden");
    splitProgress.classList.remove("hidden");
    btnSplit.disabled = true;
    try {
      const res = await fetch(API_BASE + "/api/split-devis", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ filename: filename }),
      });
      const data = await res.json().catch(function () {
        return { detail: "Réponse invalide du serveur (réponse non JSON). Vérifiez les logs du serveur." };
      });
      splitProgress.classList.add("hidden");
      btnSplit.disabled = false;
      if (!res.ok) {
        var msg = data.detail || res.statusText;
        if (msg.indexOf("Tesseract") === -1 && msg.indexOf("OCR") === -1) {
          msg = msg + " Si le problème persiste, vérifiez que Tesseract est installé.";
        }
        showSplitMessage(msg, "error");
        return;
      }
      splitResultTitle.textContent = data.devis_count + " devis créé(s) à partir de " + filename;
      var base = (data.source_stem != null ? data.source_stem : filename.replace(/\.[^.]+$/, ""));
      currentSplitSourceStem = base;
      renderSplitDownloads(base, data.devis || []);
      splitResult.classList.remove("hidden");
      loadSplitSources();
    } catch (e) {
      splitProgress.classList.add("hidden");
      btnSplit.disabled = false;
      showSplitMessage("Erreur : " + e.message, "error");
    }
  });

  function showSplitMessage(text, type) {
    splitMessage.textContent = text;
    splitMessage.className = "message " + (type === "error" ? "error" : "success");
    splitMessage.classList.remove("hidden");
  }

  if (btnMergeSplit) {
    btnMergeSplit.addEventListener("click", async function () {
      if (btnMergeSplit.disabled || !currentSplitSourceStem) return;
      var checked = splitDownloads && splitDownloads.querySelectorAll(".split-merge-cb:checked");
      if (!checked || checked.length < 2) return;
      var filenames = Array.prototype.map.call(checked, function (cb) { return cb.getAttribute("data-filename"); }).filter(Boolean);
      if (filenames.length < 2) return;
      btnMergeSplit.disabled = true;
      splitMessage.classList.add("hidden");
      try {
        var res = await fetch(API_BASE + "/api/split-devis/fusionner", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ source_stem: currentSplitSourceStem, filenames: filenames }),
        });
        var data = await res.json().catch(function () { return {}; });
        if (!res.ok) {
          var detail = data.detail;
          if (Array.isArray(detail)) detail = detail.map(function (x) { return x.msg || JSON.stringify(x); }).join(" ; ");
          else if (typeof detail !== "string") detail = detail ? String(detail) : "Erreur fusion";
          showSplitMessage(detail, "error");
          updateMergeButtonState();
          return;
        }
        renderSplitDownloads(currentSplitSourceStem, data.devis || []);
        showSplitMessage(filenames.length + " PDF fusionné(s) en un seul fichier. Les originaux ont été supprimés.", "success");
        await loadSplitSources();
        if (extractSource && extractSource.value === currentSplitSourceStem && extractSourceStem === currentSplitSourceStem) {
          runExtraction(currentSplitSourceStem);
        }
      } catch (err) {
        showSplitMessage("Erreur : " + (err.message || "fusion"), "error");
      }
      updateMergeButtonState();
    });
  }

  if (splitDownloads) {
    splitDownloads.addEventListener("click", async function (e) {
      var btnRetraiter = e.target && e.target.closest(".btn-retraiter");
      var btnSupprimer = e.target && e.target.closest(".btn-supprimer");
      if (btnSupprimer) {
        var filename = btnSupprimer.getAttribute("data-filename");
        if (!filename || !currentSplitSourceStem) return;
        btnSupprimer.disabled = true;
        splitMessage.classList.add("hidden");
        try {
          var res = await fetch(API_BASE + "/api/split-devis/supprimer", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ source_stem: currentSplitSourceStem, filename: filename }),
          });
          var data = await res.json().catch(function () { return {}; });
          if (!res.ok) {
            var detail = data.detail;
            if (Array.isArray(detail)) detail = detail.map(function (x) { return x.msg || JSON.stringify(x); }).join(" ; ");
            else if (typeof detail !== "string") detail = detail ? String(detail) : "Erreur suppression";
            showSplitMessage(detail, "error");
            btnSupprimer.disabled = false;
            return;
          }
          renderSplitDownloads(currentSplitSourceStem, data.devis || []);
          showSplitMessage("Devis supprimé.", "success");
          await loadSplitSources();
          if (extractSource && extractSource.value === currentSplitSourceStem && extractSourceStem === currentSplitSourceStem) {
            runExtraction(currentSplitSourceStem);
          }
        } catch (err) {
          showSplitMessage("Erreur : " + (err.message || "suppression"), "error");
          btnSupprimer.disabled = false;
        }
        return;
      }
      if (!btnRetraiter || btnRetraiter.disabled) return;
      var filename = btnRetraiter.getAttribute("data-filename");
      if (!filename || !currentSplitSourceStem) return;
      btnRetraiter.disabled = true;
      splitMessage.classList.add("hidden");
      try {
        var res = await fetch(API_BASE + "/api/split-devis/retraiter", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ source_stem: currentSplitSourceStem, current_filename: filename }),
        });
        var data = await res.json().catch(function () { return {}; });
        if (!res.ok) {
          var detail = data.detail;
          if (Array.isArray(detail)) detail = detail.map(function (x) { return x.msg || JSON.stringify(x); }).join(" ; ");
          else if (typeof detail !== "string") detail = detail ? String(detail) : "Erreur retraitement";
          showSplitMessage(detail, "error");
          btnRetraiter.disabled = false;
          return;
        }
        renderSplitDownloads(currentSplitSourceStem, data.devis || []);
        showSplitMessage("Page déplacée : la dernière page du PDF suivant a été mise en première position de ce PDF.", "success");
        await loadSplitSources();
        if (extractSource && extractSource.value === currentSplitSourceStem && extractSourceStem === currentSplitSourceStem) {
          runExtraction(currentSplitSourceStem);
        }
      } catch (err) {
        showSplitMessage("Erreur : " + (err.message || "retraitement"), "error");
        btnRetraiter.disabled = false;
      }
    });
  }

  function showExtractMessage(text, type) {
    var el = document.getElementById("extract-message");
    el.textContent = text;
    el.className = "message " + (type === "error" ? "error" : "success");
    el.classList.remove("hidden");
  }

  var extractSource = document.getElementById("extract-source");
  var btnExtract = document.getElementById("btn-extract");
  var extractProgress = document.getElementById("extract-progress");
  var dataTableWrap = document.getElementById("data-table-wrap");
  var dataTbody = document.getElementById("data-tbody");
  var exportSavePath = document.getElementById("export-save-path");
  var btnExportExcel = document.getElementById("btn-export-excel");
  var btnDownloadExcel = document.getElementById("btn-download-excel");
  var editModal = document.getElementById("edit-modal");
  var editForm = document.getElementById("edit-form");
  var editCancel = document.getElementById("edit-cancel");
  var editBackdrop = document.getElementById("edit-modal-backdrop");

  var extractedRecords = [];
  var extractSourceStem = "";

  var EDIT_COLUMN_TO_RECORD = {
    code_fiche: "code_fiche",
    raison_sociale_demandeur: "nom_site_beneficiaire",
    date_rai: "date_rai",
    prime_cee: "prime_cee",
    nom_beneficiaire: "nom_beneficiaire",
    prenom_beneficiaire: "prenom_beneficiaire",
    adresse_op: "adresse_des_travaux",
    cp_op: "code_postal_travaux",
    ville_op: "ville_travaux",
    tel: "tel",
    mail: "mail",
    nom_site: "nom_site_beneficiaire",
    adresse_site: "adresse_client",
    cp_site: "code_postal",
    ville_site: "ville",
    num_devis: "num_devis",
    num_client: "num_client",
    montant_ttc: "prime_cee",
    nb_luminaires: "nombre_luminaires",
    puissance_led: "puissance_totale_led_w",
    irc_inf90: "irc",
    precision_secteur: "secteur_activite"
  };

  function sirenFromSiret(s) {
    if (s == null || s === "") return "";
    var digits = String(s).replace(/\D/g, "");
    return digits.length ? digits.substring(0, 9) : "";
  }

  function numVal(v) {
    if (v == null || v === "") return "";
    var n = parseFloat(String(v).replace(/\s/g, "").replace(",", "."));
    return isNaN(n) ? "" : n;
  }

  function datePlus15(dateStr) {
    if (!dateStr || !String(dateStr).trim()) return "";
    var s = String(dateStr).trim();
    var m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2,4})$/);
    if (!m) return "";
    var d = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1, y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    var dt = new Date(y, mo, d);
    dt.setDate(dt.getDate() + 15);
    var dd = dt.getDate(), mm = dt.getMonth() + 1;
    return (dd < 10 ? "0" : "") + dd + "/" + (mm < 10 ? "0" : "") + mm + "/" + dt.getFullYear();
  }

  function telDigits(s) {
    if (s == null || s === "") return "";
    return String(s).replace(/\D/g, "");
  }

  function dateRaiFallback(index) {
    var base = new Date(2025, 10, 13);
    var end = new Date(2025, 11, 28);
    var dayOffset = Math.floor((index || 0) / 4);
    var d = new Date(base);
    d.setDate(d.getDate() + dayOffset);
    if (d > end) d = end;
    var dd = d.getDate(), mm = d.getMonth() + 1;
    return (dd < 10 ? "0" : "") + dd + "/" + (mm < 10 ? "0" : "") + mm + "/" + d.getFullYear();
  }

  function getExcelStyleValue(rec, index, key) {
    var r = rec || {};
    switch (key) {
      case "opération": return "";
      case "type": return "TPM";
      case "code_fiche": return (r.code_fiche || "BAT-EQ-127").trim();
      case "raison_sociale_demandeur": return (r.nom_site_beneficiaire || "").trim();
      case "siren_demandeur": return sirenFromSiret(r.siret);
      case "reference_pixel": return "";
      case "date_rai":
        var drai = (r.date_rai || "").trim();
        return drai ? drai : dateRaiFallback(index);
      case "prime_cee": return (r.prime_cee || "").trim();
      case "date_engagement":
        var drai2 = (r.date_rai || "").trim();
        return datePlus15(drai2 || dateRaiFallback(index));
      case "mandataire": return "OTC FLOW France";
      case "siren_mandataire": return "953658036";
      case "nature_bonification": return "ZNI";
      case "nom_beneficiaire": return (r.nom_beneficiaire || "").trim();
      case "prenom_beneficiaire": return (r.prenom_beneficiaire || "").trim();
      case "adresse_op": return (r.adresse_des_travaux || "").trim();
      case "cp_op": return (r.code_postal_travaux || r.code_postal || "").trim().replace(/\s/g, "");
      case "ville_op": return (r.ville_travaux || r.ville || "").trim();
      case "tel": return telDigits(r.tel);
      case "mail": return (r.mail || "").trim();
      case "nom_site": return (r.nom_site_beneficiaire || "").trim();
      case "adresse_site": return (r.adresse_client || "").trim();
      case "cp_site": return (r.code_postal || "").trim().replace(/\s/g, "");
      case "ville_site": return (r.ville || "").trim();
      case "tel_site": return telDigits(r.tel);
      case "mail_site": return (r.mail || "").trim();
      case "siren_pro": return "421747007";
      case "raison_pro": return "LECBA.TP SARL";
      case "num_devis": return (r.num_devis || r.num_client || "").trim();
      case "num_client": return (r.num_client || "").trim();
      case "montant_ttc": return (r.prime_cee || "").trim();
      case "raison_pro_devis": return "LECBA.TP SARL";
      case "siren_pro_devis": return "42174700700035";
      case "nb_luminaires": return (r.nombre_luminaires || "").trim();
      case "puissance_led":
        var nb = parseFloat(String(r.nombre_luminaires || "").replace(/\s/g, "").replace(",", "."));
        return isNaN(nb) ? "" : String(Math.round(nb * 250));
      case "irc_inf90": return (r.irc || "").trim();
      case "irc_sup90": return "";
      case "secteur": return "";
      case "precision_secteur": return (r.secteur_activite || "").trim();
      default: return "";
    }
  }

  var TABLE_COLUMNS = [
    { label: "Opération n°", key: "opération", get: function (rec, i) { return getExcelStyleValue(rec, i, "opération"); } },
    { label: "Type de trame (TOP/TPM)", key: "type", get: function (rec, i) { return getExcelStyleValue(rec, i, "type"); } },
    { label: "Code Fiche", key: "code_fiche", get: function (rec, i) { return getExcelStyleValue(rec, i, "code_fiche"); } },
    { label: "RAISON SOCIALE du demandeur", key: "raison_sociale_demandeur", get: function (rec, i) { return getExcelStyleValue(rec, i, "raison_sociale_demandeur"); } },
    { label: "SIREN du demandeur", key: "siren_demandeur", get: function (rec, i) { return getExcelStyleValue(rec, i, "siren_demandeur"); } },
    { label: "REFERENCE PIXEL", key: "reference_pixel", get: function (rec, i) { return getExcelStyleValue(rec, i, "reference_pixel"); } },
    { label: "DATE d'envoi du RAI", key: "date_rai", get: function (rec, i) { return getExcelStyleValue(rec, i, "date_rai"); } },
    { label: "MONTANT incitation CEE", key: "prime_cee", get: function (rec, i) { return getExcelStyleValue(rec, i, "prime_cee"); } },
    { label: "DATE D'ENGAGEMENT", key: "date_engagement", get: function (rec, i) { return getExcelStyleValue(rec, i, "date_engagement"); } },
    { label: "Raison sociale du mandataire", key: "mandataire", get: function (rec, i) { return getExcelStyleValue(rec, i, "mandataire"); } },
    { label: "SIREN du mandataire", key: "siren_mandataire", get: function (rec, i) { return getExcelStyleValue(rec, i, "siren_mandataire"); } },
    { label: "Nature de la bonification", key: "nature_bonification", get: function (rec, i) { return getExcelStyleValue(rec, i, "nature_bonification"); } },
    { label: "NOM du bénéficiaire", key: "nom_beneficiaire", get: function (rec, i) { return getExcelStyleValue(rec, i, "nom_beneficiaire"); } },
    { label: "PRENOM du bénéficiaire", key: "prenom_beneficiaire", get: function (rec, i) { return getExcelStyleValue(rec, i, "prenom_beneficiaire"); } },
    { label: "ADRESSE de l'opération", key: "adresse_op", get: function (rec, i) { return getExcelStyleValue(rec, i, "adresse_op"); } },
    { label: "CODE POSTAL", key: "cp_op", get: function (rec, i) { return getExcelStyleValue(rec, i, "cp_op"); } },
    { label: "VILLE", key: "ville_op", get: function (rec, i) { return getExcelStyleValue(rec, i, "ville_op"); } },
    { label: "Tél bénéficiaire", key: "tel", get: function (rec, i) { return getExcelStyleValue(rec, i, "tel"); } },
    { label: "Mail bénéficiaire", key: "mail", get: function (rec, i) { return getExcelStyleValue(rec, i, "mail"); } },
    { label: "NOM DU SITE bénéficiaire", key: "nom_site", get: function (rec, i) { return getExcelStyleValue(rec, i, "nom_site"); } },
    { label: "ADRESSE opération (site)", key: "adresse_site", get: function (rec, i) { return getExcelStyleValue(rec, i, "adresse_site"); } },
    { label: "CODE POSTAL (site)", key: "cp_site", get: function (rec, i) { return getExcelStyleValue(rec, i, "cp_site"); } },
    { label: "VILLE (site)", key: "ville_site", get: function (rec, i) { return getExcelStyleValue(rec, i, "ville_site"); } },
    { label: "Tél (site)", key: "tel_site", get: function (rec, i) { return getExcelStyleValue(rec, i, "tel_site"); } },
    { label: "Mail (site)", key: "mail_site", get: function (rec, i) { return getExcelStyleValue(rec, i, "mail_site"); } },
    { label: "SIREN du professionnel", key: "siren_pro", get: function (rec, i) { return getExcelStyleValue(rec, i, "siren_pro"); } },
    { label: "RAISON SOCIALE du professionnel", key: "raison_pro", get: function (rec, i) { return getExcelStyleValue(rec, i, "raison_pro"); } },
    { label: "NUMERO de devis", key: "num_devis", get: function (rec, i) { return getExcelStyleValue(rec, i, "num_devis"); } },
    { label: "Numéro Client", key: "num_client", get: function (rec, i) { return getExcelStyleValue(rec, i, "num_client"); } },
    { label: "MONTANT du devis (€ TTC)", key: "montant_ttc", get: function (rec, i) { return getExcelStyleValue(rec, i, "montant_ttc"); } },
    { label: "RAISON SOCIALE pro sur devis", key: "raison_pro_devis", get: function (rec, i) { return getExcelStyleValue(rec, i, "raison_pro_devis"); } },
    { label: "SIREN pro sur devis", key: "siren_pro_devis", get: function (rec, i) { return getExcelStyleValue(rec, i, "siren_pro_devis"); } },
    { label: "Nombre luminaires", key: "nb_luminaires", get: function (rec, i) { return getExcelStyleValue(rec, i, "nb_luminaires"); } },
    { label: "Puissance totale LED", key: "puissance_led", get: function (rec, i) { return getExcelStyleValue(rec, i, "puissance_led"); } },
    { label: "Puissance IRC < 90", key: "irc_inf90", get: function (rec, i) { return getExcelStyleValue(rec, i, "irc_inf90"); } },
    { label: "Puissance IRC ≥ 90", key: "irc_sup90", get: function (rec, i) { return getExcelStyleValue(rec, i, "irc_sup90"); } },
    { label: "Secteur concerné", key: "secteur", get: function (rec, i) { return getExcelStyleValue(rec, i, "secteur"); } },
    { label: "Précision secteur", key: "precision_secteur", get: function (rec, i) { return getExcelStyleValue(rec, i, "precision_secteur"); } }
  ];

  function renderEditFormFields() {
    var container = document.getElementById("edit-form-fields");
    if (!container) return;
    container.innerHTML = TABLE_COLUMNS.map(function (col) {
      var id = "edit-" + col.key;
      return "<label>" + escapeHtml(col.label) + " <input type='text' id='" + id + "' /></label>";
    }).join("");
  }

  function renderDataTableHeaders() {
    var headRow = document.getElementById("data-table-head-row");
    if (!headRow) return;
    headRow.innerHTML = TABLE_COLUMNS.map(function (col) {
      return "<th>" + escapeHtml(col.label) + "</th>";
    }).join("") + "<th></th>";
  }

  async function loadSplitSources() {
    try {
      var selected = extractSource && extractSource.value;
      var res = await fetch(API_BASE + "/api/splits/sources");
      var data = await res.json();
      var sources = data.sources || [];
      var sel = extractSource;
      var first = sel.querySelector('option[value=""]');
      sel.innerHTML = "";
      if (first) sel.appendChild(first);
      sources.forEach(function (s) {
        var o = document.createElement("option");
        o.value = s.source_stem;
        o.textContent = s.source_stem + " (" + s.file_count + " devis)";
        sel.appendChild(o);
      });
      if (selected && sources.some(function (s) { return s.source_stem === selected; })) {
        sel.value = selected;
      }
      btnExtract.disabled = !sel.value;
    } catch (_) {}
  }

  extractSource.addEventListener("change", function () {
    btnExtract.disabled = !extractSource.value;
  });

  function renderDataTable() {
    var nbLuminairesColIndex = TABLE_COLUMNS.findIndex(function (c) { return c.key === "nb_luminaires"; });
    dataTbody.innerHTML = extractedRecords
      .map(function (rec, i) {
        var cells = TABLE_COLUMNS.map(function (col, colIdx) {
          var val = col.get(rec, i);
          var content = escapeHtml(val != null ? String(val) : "");
          if (colIdx === nbLuminairesColIndex) {
            return "<td class='cell-nb-luminaires' data-index='" + i + "' title='Cliquer pour voir le détail de la somme'>" + content + "</td>";
          }
          return "<td>" + content + "</td>";
        });
        cells.push("<td class='btn-cell'><button type='button' class='btn btn-ghost btn-sm btn-edit'>Modifier</button></td>");
        return "<tr data-index='" + i + "'>" + cells.join("") + "</tr>";
      })
      .join("");
    dataTableWrap.classList.remove("hidden");
    dataTbody.querySelectorAll(".btn-edit").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var tr = btn.closest("tr");
        var i = parseInt(tr.getAttribute("data-index"), 10);
        openEditModal(i);
      });
    });
    dataTbody.querySelectorAll(".cell-nb-luminaires").forEach(function (td) {
      td.style.cursor = "pointer";
      td.addEventListener("click", function (e) {
        e.stopPropagation();
        var i = parseInt(td.getAttribute("data-index"), 10);
        var rec = extractedRecords[i];
        var detail = (rec && rec.nombre_luminaires_detail) ? rec.nombre_luminaires_detail.slice() : [];
        var sum = (rec && rec.nombre_luminaires) ? String(rec.nombre_luminaires).trim() : "";
        showNombreLuminairesDetail(detail, sum, td, i);
      });
    });
  }

  function showNombreLuminairesDetail(detail, sum, anchorEl, recordIndex) {
    var existing = document.getElementById("popover-detail-somme");
    if (existing) existing.remove();
    var rec = extractedRecords[recordIndex];
    var pop = document.createElement("div");
    pop.id = "popover-detail-somme";
    pop.className = "popover-detail-somme popover-detail-somme-editable";
    var values = detail.length ? detail.slice() : (sum ? [sum] : []);
    if (values.length === 0) values.push("");
    var inputsHtml = values.map(function (v, idx) {
      return "<div class='popover-detail-row'><label>Valeur " + (idx + 1) + "</label><input type='text' class='popover-detail-input' value='" + escapeHtml(String(v)) + "' placeholder='ex. 313' /></div>";
    }).join("");
    pop.innerHTML = "<div class='popover-detail-somme-inner'>" +
      "<strong>Détail de la somme (Nombre luminaires)</strong>" +
      "<div class='popover-detail-inputs'>" + inputsHtml + "</div>" +
      "<button type='button' class='btn btn-ghost btn-sm popover-detail-add'>+ Ajouter une valeur</button>" +
      "<p class='popover-detail-total'>Total : <span class='popover-detail-total-value'>" + escapeHtml(sum || "—") + "</span></p>" +
      "<div class='popover-detail-actions'>" +
      "<button type='button' class='btn btn-primary btn-sm popover-detail-save'>Enregistrer</button> " +
      "<button type='button' class='btn btn-ghost btn-sm popover-detail-close'>Fermer</button>" +
      "</div></div>";
    document.body.appendChild(pop);
    var rect = anchorEl.getBoundingClientRect();
    pop.style.left = Math.min(rect.left, window.innerWidth - 340) + "px";
    pop.style.top = (rect.bottom + 6) + "px";

    function refreshTotalPreview() {
      var inputs = pop.querySelectorAll(".popover-detail-input");
      var total = 0;
      var valid = [];
      inputs.forEach(function (inp) {
        var raw = (inp.value || "").trim().replace(",", ".");
        if (raw === "") return;
        var n = parseFloat(raw);
        if (!isNaN(n) && n >= 0) {
          total += n;
          valid.push(n === Math.floor(n) ? String(Math.floor(n)) : n.toFixed(2).replace(".", ","));
        }
      });
      var sumStr = total === Math.floor(total) ? String(Math.floor(total)) : total.toFixed(2).replace(".", ",");
      var totalEl = pop.querySelector(".popover-detail-total-value");
      if (totalEl) totalEl.textContent = valid.length ? sumStr : "—";
    }

    pop.querySelectorAll(".popover-detail-input").forEach(function (inp) {
      inp.addEventListener("input", refreshTotalPreview);
      inp.addEventListener("change", refreshTotalPreview);
    });

    pop.querySelector(".popover-detail-add").addEventListener("click", function () {
      var container = pop.querySelector(".popover-detail-inputs");
      var n = container.querySelectorAll(".popover-detail-row").length + 1;
      var div = document.createElement("div");
      div.className = "popover-detail-row";
      div.innerHTML = "<label>Valeur " + n + "</label><input type='text' class='popover-detail-input' placeholder='ex. 134' />";
      container.appendChild(div);
      div.querySelector("input").addEventListener("input", refreshTotalPreview);
      div.querySelector("input").addEventListener("change", refreshTotalPreview);
    });

    pop.querySelector(".popover-detail-save").addEventListener("click", function () {
      var inputs = pop.querySelectorAll(".popover-detail-input");
      var detailNew = [];
      var total = 0;
      inputs.forEach(function (inp) {
        var raw = (inp.value || "").trim().replace(",", ".");
        if (raw === "") return;
        var n = parseFloat(raw);
        if (!isNaN(n) && n >= 0) {
          total += n;
          detailNew.push(n === Math.floor(n) ? String(Math.floor(n)) : n.toFixed(2).replace(".", ","));
        }
      });
      var sumStr = total === Math.floor(total) ? String(Math.floor(total)) : (total.toFixed(2).replace(".", ","));
      if (rec) {
        rec.nombre_luminaires = detailNew.length ? sumStr : "";
        rec.nombre_luminaires_detail = detailNew;
      }
      pop.remove();
      renderDataTable();
      exportAndSave();
    });

    pop.querySelector(".popover-detail-close").addEventListener("click", function () { pop.remove(); });
    function closePop(ev) {
      if (pop.parentNode && !pop.contains(ev.target) && ev.target !== anchorEl) {
        pop.remove();
        document.removeEventListener("click", closePop);
      }
    }
    document.addEventListener("click", closePop);
  }

  function openEditModal(index) {
    var rec = extractedRecords[index];
    if (!rec) return;
    document.getElementById("edit-index").value = index;
    TABLE_COLUMNS.forEach(function (col) {
      var el = document.getElementById("edit-" + col.key);
      if (el) {
        var val = col.get(rec, index);
        el.value = (val != null ? String(val) : "") || "";
      }
    });
    var iframe = document.getElementById("edit-pdf-iframe");
    var placeholder = document.getElementById("edit-pdf-placeholder");
    var pdfFilename = rec._filename || (extractSourceStem && rec.num_devis ? extractSourceStem + "_" + rec.num_devis + ".pdf" : null);
    if (pdfFilename && extractSourceStem) {
      var pdfUrl = API_BASE + "/api/splits/" + encodeURIComponent(extractSourceStem) + "/" + encodeURIComponent(pdfFilename) + "?display=1";
      iframe.src = pdfUrl;
      iframe.classList.remove("hidden");
      placeholder.classList.add("hidden");
    } else {
      iframe.src = "";
      iframe.classList.add("hidden");
      placeholder.classList.remove("hidden");
    }
    editModal.classList.remove("hidden");
  }

  function closeEditModal() {
    editModal.classList.add("hidden");
    var iframe = document.getElementById("edit-pdf-iframe");
    iframe.src = "";
  }

  function highlightModifiedRow(index) {
    if (!dataTbody) return;
    var tr = dataTbody.querySelector("tr[data-index='" + index + "']");
    if (!tr) return;
    tr.classList.remove("row-modified");
    tr.classList.add("row-modified");
    tr.scrollIntoView({ behavior: "smooth", block: "nearest" });
    setTimeout(function () {
      tr.classList.remove("row-modified");
    }, 4000);
  }

  editForm.addEventListener("submit", function (e) {
    e.preventDefault();
    var i = parseInt(document.getElementById("edit-index").value, 10);
    if (isNaN(i) || i < 0 || i >= extractedRecords.length) return;
    var rec = extractedRecords[i];
    if (rec) {
      TABLE_COLUMNS.forEach(function (col) {
        var el = document.getElementById("edit-" + col.key);
        var recKey = EDIT_COLUMN_TO_RECORD[col.key];
        if (el && recKey) {
          var v = (el.value != null ? String(el.value).trim() : "");
          rec[recKey] = v || null;
        }
      });
    }
    closeEditModal();
    renderDataTable();
    highlightModifiedRow(i);
    exportAndSave();
  });

  editCancel.addEventListener("click", function () {
    var idx = parseInt(document.getElementById("edit-index").value, 10);
    closeEditModal();
    highlightModifiedRow(idx);
  });
  editBackdrop.addEventListener("click", function () {
    var idx = parseInt(document.getElementById("edit-index").value, 10);
    closeEditModal();
    highlightModifiedRow(idx);
  });

  async function runExtraction(source) {
    if (!source) return;
    document.getElementById("extract-message").classList.add("hidden");
    extractProgress.classList.remove("hidden");
    btnExtract.disabled = true;
    try {
      var res = await fetch(API_BASE + "/api/extract-devis", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ source_stem: source }),
      });
      var data = await res.json().catch(function () { return {}; });
      extractProgress.classList.add("hidden");
      btnExtract.disabled = false;
      if (!res.ok) {
        showExtractMessage(data.detail || "Erreur extraction", "error");
        return;
      }
      var results = data.results || [];
      extractSourceStem = source;
      var allKeys = ["num_devis", "num_client", "nom_client", "siret", "adresse_client", "code_postal", "ville", "tel", "mail", "represente_par", "nom_beneficiaire", "prenom_beneficiaire", "secteur_activite", "adresse_des_travaux", "code_postal_travaux", "ville_travaux", "nom_site_beneficiaire", "date_rai", "code_fiche", "prime_cee", "total_ht", "total_ttc", "raison_sociale_professionnel", "siren_professionnel", "nombre_luminaires", "puissance_totale_led_w", "irc"];
      extractedRecords = results.filter(function (r) { return !r.error; }).map(function (r) {
        var o = {};
        allKeys.forEach(function (id) {
          o[id] = r[id] != null ? String(r[id]) : "";
        });
        if (r._filename) o._filename = r._filename;
        if (r.nombre_luminaires_detail && Array.isArray(r.nombre_luminaires_detail)) o.nombre_luminaires_detail = r.nombre_luminaires_detail;
        return o;
      });
      if (results.length && results.some(function (r) { return r.error; })) {
        showExtractMessage(extractedRecords.length + " ligne(s) extraite(s). Certains fichiers ont eu des erreurs.", "error");
      } else {
        showExtractMessage(extractedRecords.length + " devis extrait(s).", "success");
      }
      renderDataTable();
      exportAndSave();
    } catch (e) {
      extractProgress.classList.add("hidden");
      btnExtract.disabled = false;
      showExtractMessage("Erreur : " + e.message, "error");
    }
  }

  btnExtract.addEventListener("click", function () {
    runExtraction(extractSource.value);
  });

  function getExportPath() {
    return (exportSavePath && exportSavePath.value) ? exportSavePath.value.trim() : "";
  }

  function exportAndSave() {
    var path = getExportPath();
    if (!extractedRecords.length) {
      showExtractMessage("Aucune donnée à exporter. Lancez d'abord l'extraction.", "error");
      return;
    }
    if (!path) {
      showExtractMessage("Indiquez un chemin d'enregistrement (champ au-dessus du bouton).", "error");
      return;
    }
    if (btnExportExcel) btnExportExcel.disabled = true;
    fetch(API_BASE + "/api/export-excel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ records: extractedRecords, save_path: path }),
    })
      .then(function (res) {
        return res.text().then(function (text) {
          var data = {};
          try {
            data = text ? JSON.parse(text) : {};
          } catch (_) {}
          if (res.ok && data.saved) {
            showExtractMessage("Excel enregistré : " + (data.path || path), "success");
          } else {
            var msg = data.detail || (res.ok ? "Erreur enregistrement" : "Erreur " + res.status);
            if (!res.ok && text && text.length < 500 && !data.detail) msg = text;
            showExtractMessage(msg, "error");
          }
        });
      })
      .catch(function (e) {
        showExtractMessage("Erreur : " + (e.message || "enregistrement Excel"), "error");
      })
      .finally(function () {
        if (btnExportExcel) btnExportExcel.disabled = false;
      });
  }

  if (btnExportExcel) {
    btnExportExcel.addEventListener("click", function () {
      exportAndSave();
    });
  }

  if (btnDownloadExcel) {
    btnDownloadExcel.addEventListener("click", async function () {
      if (!extractedRecords.length) {
        showExtractMessage("Aucune donnée à exporter. Lancez d’abord l’extraction.", "error");
        return;
      }
      try {
        var res = await fetch(API_BASE + "/api/export-excel", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ records: extractedRecords }),
        });
        if (!res.ok) {
          var text = await res.text();
          var detail = "Erreur export (" + res.status + ")";
          try {
            var err = JSON.parse(text);
            if (err && err.detail != null) detail = Array.isArray(err.detail) ? err.detail.map(function (x) { return x.msg || JSON.stringify(x); }).join(" ; ") : (err.detail || detail);
          } catch (_) {
            if (text && text.length < 300) detail = text;
          }
          showExtractMessage(detail, "error");
          return;
        }
        var blob = await res.blob();
        var a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "recensement_bat_eq_127.xlsx";
        a.click();
        URL.revokeObjectURL(a.href);
        showExtractMessage("Fichier téléchargé.", "success");
      } catch (e) {
        showExtractMessage("Erreur : " + e.message, "error");
      }
    });
  }

  async function loadEnregistrements() {
    if (!enregistrementsUploads || !enregistrementsExports) return;
    try {
      var res = await fetch(API_BASE + "/api/enregistrements");
      var data = await res.json().catch(function () { return { uploads: [], exports: [] }; });
      var uploads = data.uploads || [];
      var exports = data.exports || [];
      var base = API_BASE + "/api/enregistrements/download?path=";
      enregistrementsUploads.innerHTML = uploads.map(function (f) {
        var url = base + encodeURIComponent(f.path);
        return "<li><a href='" + url + "' download>" + escapeHtml(f.name) + "</a> <span class='size'>(" + formatSize(f.size) + ")</span></li>";
      }).join("");
      enregistrementsExports.innerHTML = exports.map(function (f) {
        var url = base + encodeURIComponent(f.path);
        return "<li><a href='" + url + "' download>" + escapeHtml(f.name) + "</a> <span class='size'>(" + formatSize(f.size) + ")</span></li>";
      }).join("");
      if (uploadsReadingEmpty) uploadsReadingEmpty.classList.toggle("hidden", uploads.length > 0);
      if (exportsReadingEmpty) exportsReadingEmpty.classList.toggle("hidden", exports.length > 0);
    } catch (e) {
      if (uploadsReadingEmpty) uploadsReadingEmpty.classList.remove("hidden");
      if (exportsReadingEmpty) exportsReadingEmpty.classList.remove("hidden");
    }
  }

  if (btnLoadEnregistrements) btnLoadEnregistrements.addEventListener("click", loadEnregistrements);

  btnRefresh.addEventListener("click", function () {
    loadUploadedList();
    loadSplitSources();
    loadEnregistrements();
  });

  renderDataTableHeaders();
  renderEditFormFields();
  loadUploadedList();
  loadSplitSources();
  loadEnregistrements();
})();