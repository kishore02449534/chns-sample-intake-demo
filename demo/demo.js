const SHEET_COLUMNS = [
    { key: "bfiCode", label: "BFI_Code" },
    { key: "tubeLabel", label: "Tube Label" },
    { key: "receivedDate", label: "Received Date" },
    { key: "category", label: "Category" },
    { key: "species", label: "Species" },
    { key: "cut", label: "Cut / Type" },
    { key: "breed", label: "Breed / Cultivar" },
    { key: "practice", label: "Practice" },
    { key: "farmName", label: "Farm Name" },
];

const CHECK_EVERY_MINUTES = 60;

const demoState = {
    skippedMinutes: 0,
    sheetRows: [],
    savedRows: [],
    unsavedChanges: false,
    fileVersion: 1,
    fileSavedAt: null,
    lastSeenVersion: 1,
    existingCodes: new Set(),
    waitingSamples: [],
    liveSamples: [],
    farmCards: [],
    heldRows: [],
    latestCheck: null,
    nextCheckAt: null,
    history: [],
    nextRowId: 1,
    nextCardId: 1,
    goodExampleIndex: 0,
    problemExampleIndex: 0,
    openCardId: null,
};

function compactKey(text) {
    return String(text || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function codeKey(code) {
    return String(code || "").replace(/\s+/g, "").toUpperCase();
}

function escapeHtml(text) {
    return String(text ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function currentTime() {
    return addMinutes(new Date(), demoState.skippedMinutes);
}

function addMinutes(when, minutes) {
    return new Date(when.getTime() + minutes * 60000);
}

function shortDate(when) {
    const month = String(when.getMonth() + 1).padStart(2, "0");
    const day = String(when.getDate()).padStart(2, "0");
    return `${month}/${day}/${when.getFullYear()}`;
}

function shortDateTime(when) {
    const clockTime = when.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
    return `${shortDate(when)}, ${clockTime}`;
}

function samplesWord(count) {
    return count === 1 ? "1 sample" : `${count} samples`;
}

function rowsWord(count) {
    return count === 1 ? "1 row" : `${count} rows`;
}

function readDate(text) {
    const trimmed = String(text || "").trim();
    const usMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
    const isoMatch = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    let year;
    let month;
    let day;
    if (usMatch) {
        [, month, day, year] = usMatch.map(Number);
        if (year < 100) year += 2000;
    } else if (isoMatch) {
        [, year, month, day] = isoMatch.map(Number);
    } else {
        return null;
    }
    const parsed = new Date(year, month - 1, day);
    if (parsed.getMonth() !== month - 1 || parsed.getDate() !== day) return null;
    return parsed;
}

function makeSheetRow(cellValues) {
    const cells = {};
    SHEET_COLUMNS.forEach((column, columnIndex) => {
        cells[column.key] = cellValues[columnIndex] || "";
    });
    const sheetRow = { rowId: demoState.nextRowId, cells };
    demoState.nextRowId += 1;
    return sheetRow;
}

function copyRows(rows) {
    return rows.map((sheetRow) => ({ rowId: sheetRow.rowId, cells: { ...sheetRow.cells } }));
}

function findFarm(sheetFarmName) {
    const farm = DEMO_REFERENCE.farms.find((knownFarm) => compactKey(knownFarm.farmName) === compactKey(sheetFarmName));
    if (!farm) throw new Error(`Farm '${sheetFarmName}' was not found. Add it on the Manage Farms page first.`);
    return farm;
}

function findSampleType(sheetCategory) {
    const sampleType = DEMO_REFERENCE.sampleTypeByCategory[compactKey(sheetCategory)];
    if (!sampleType) throw new Error(`Category '${sheetCategory}' is not in the category list.`);
    return sampleType;
}

function bracketWord(foodName) {
    const bracketMatch = foodName.match(/\(([^)]*)\)/);
    return bracketMatch ? compactKey(bracketMatch[1]) : "";
}

function findFood(sheetCategory, sheetSpecies) {
    const speciesWord = compactKey(sheetSpecies);
    const aliasName =
        DEMO_REFERENCE.speciesByCategoryWord[`${compactKey(sheetCategory)}:${speciesWord}`] ||
        DEMO_REFERENCE.speciesByCategoryWord[`:${speciesWord}`];
    const foods = DEMO_REFERENCE.foods;
    const food =
        foods.find((knownFood) => knownFood.foodName === aliasName) ||
        foods.find((knownFood) => compactKey(knownFood.foodName) === speciesWord) ||
        foods.find((knownFood) => bracketWord(knownFood.foodName) === speciesWord);
    if (!food) throw new Error(`Species '${sheetSpecies}' was not found.`);
    return food;
}

function findBreed(food, sheetBreed) {
    const breedName = sheetBreed.trim() || DEMO_REFERENCE.breedWhenBlank;
    const breed = food.breeds.find((knownBreed) => compactKey(knownBreed) === compactKey(breedName));
    if (!breed) throw new Error(`Breed / Cultivar '${breedName}' was not found for ${food.foodName}.`);
    return breed;
}

function findCut(sheetCut) {
    const cutWord = compactKey(sheetCut);
    const cutName = !cutWord ? DEMO_REFERENCE.cutWhenBlank : DEMO_REFERENCE.cutBySheetWord[cutWord] || sheetCut.trim();
    const cut = DEMO_REFERENCE.cuts.find((knownCut) => compactKey(knownCut) === compactKey(cutName));
    if (!cut) throw new Error(`Cut / Type '${sheetCut}' was not found.`);
    return cut;
}

function findPractices(sheetPractice) {
    const practiceWords = String(sheetPractice || "")
        .split(/[,;]/)
        .map((word) => word.trim())
        .filter(Boolean);
    const missingWords = [];
    const practiceNames = [];
    practiceWords.forEach((word) => {
        const practice = DEMO_REFERENCE.practices.find((knownPractice) => compactKey(knownPractice) === compactKey(word));
        if (practice) practiceNames.push(practice);
        else missingWords.push(`'${word}'`);
    });
    if (missingWords.length > 0) throw new Error(`Practice ${missingWords.join(", ")} was not found.`);
    return practiceNames;
}

function tryMatch(reasons, matcher) {
    try {
        return matcher();
    } catch (matchError) {
        reasons.push(matchError.message);
        return null;
    }
}

function checkOneRow(cells) {
    const reasons = [];
    const requiredCells = [
        ["tubeLabel", "Tube Label is empty."],
        ["farmName", "Farm Name is empty."],
        ["category", "Category is empty."],
        ["species", "Species is empty."],
    ];
    requiredCells.forEach(([cellKey, message]) => {
        if (!cells[cellKey].trim()) reasons.push(message);
    });

    const receivedText = cells.receivedDate.trim();
    const receivedDate = readDate(receivedText);
    if (!receivedText) reasons.push("Received Date is empty.");
    else if (!receivedDate) reasons.push(`Received Date '${receivedText}' is not a date.`);

    const farm = cells.farmName.trim() ? tryMatch(reasons, () => findFarm(cells.farmName)) : null;
    const sampleType = cells.category.trim() ? tryMatch(reasons, () => findSampleType(cells.category)) : null;
    const food = cells.species.trim() ? tryMatch(reasons, () => findFood(cells.category, cells.species)) : null;
    const breed = food ? tryMatch(reasons, () => findBreed(food, cells.breed)) : null;
    const cut = tryMatch(reasons, () => findCut(cells.cut));
    const practices = tryMatch(reasons, () => findPractices(cells.practice));

    if (food && breed && cut && !food.items.includes(`${breed}|${cut}`)) {
        reasons.push(`${food.foodName}, ${breed}, ${cut} is not set up as a food item yet.`);
    }
    if (reasons.length > 0) return { reasons };

    return {
        sample: {
            bfiCode: codeKey(cells.bfiCode),
            tubeLabel: cells.tubeLabel.trim(),
            farmName: farm.farmName,
            farmLocation: farm.location,
            sampleType,
            foodDescription: `${food.foodName}, ${breed}, ${cut}`,
            practices,
            receivedDate,
        },
    };
}

function looksLikeSample(cells) {
    return ["bfiCode", "tubeLabel", "farmName", "category", "species"].some((cellKey) => cells[cellKey].trim());
}

function codesInUse() {
    const liveCodes = new Set(demoState.liveSamples.map((sample) => sample.bfiCode));
    const waitingCodes = new Set(demoState.waitingSamples.map((sample) => sample.bfiCode));
    return { liveCodes, waitingCodes };
}

function checkSavedFile() {
    const { liveCodes, waitingCodes } = codesInUse();
    const sampleRows = demoState.savedRows
        .map((sheetRow, rowIndex) => ({ sheetRowNumber: rowIndex + 2, cells: sheetRow.cells }))
        .filter((sheetRow) => looksLikeSample(sheetRow.cells));

    const rowsByCode = new Map();
    sampleRows.forEach((sheetRow) => {
        const key = codeKey(sheetRow.cells.bfiCode);
        if (!key) return;
        if (!rowsByCode.has(key)) rowsByCode.set(key, []);
        rowsByCode.get(key).push(sheetRow.sheetRowNumber);
    });

    const outcome = { oldRows: 0, alreadyWaiting: 0, alreadyLive: 0, ready: [], held: [] };
    sampleRows.forEach((sheetRow) => {
        const key = codeKey(sheetRow.cells.bfiCode);
        const heldRow = (reasons) => ({ sheetRowNumber: sheetRow.sheetRowNumber, bfiCode: key, reasons });
        if (!key) {
            outcome.held.push(heldRow(["BFI_Code is empty."]));
        } else if (demoState.existingCodes.has(key)) {
            outcome.oldRows += 1;
        } else if (liveCodes.has(key)) {
            outcome.alreadyLive += 1;
        } else if (waitingCodes.has(key)) {
            outcome.alreadyWaiting += 1;
        } else if (rowsByCode.get(key).length > 1) {
            const rowList = rowsByCode.get(key).join(", ");
            outcome.held.push(heldRow([`BFI_Code ${key} appears more than once in the file (rows ${rowList}).`]));
        } else {
            const rowResult = checkOneRow(sheetRow.cells);
            if (rowResult.sample) outcome.ready.push(rowResult.sample);
            else outcome.held.push(heldRow(rowResult.reasons));
        }
    });
    return outcome;
}

function pendingCardForFarm(farmName) {
    const existingCard = demoState.farmCards.find((farmCard) => farmCard.farmName === farmName && farmCard.status === "waiting");
    if (existingCard) return existingCard;
    const newCard = { cardId: demoState.nextCardId, farmName, status: "waiting", createdAt: currentTime() };
    demoState.nextCardId += 1;
    demoState.farmCards.push(newCard);
    return newCard;
}

function addToInbox(readySamples) {
    const farmsTouched = new Map();
    readySamples.forEach((sample) => {
        const farmCard = pendingCardForFarm(sample.farmName);
        demoState.waitingSamples.push({ ...sample, cardId: farmCard.cardId });
        farmsTouched.set(sample.farmName, (farmsTouched.get(sample.farmName) || 0) + 1);
    });
    return farmsTouched;
}

function runCheck(checkKind) {
    const checkedAt = currentTime();
    const fileChanged = demoState.fileVersion !== demoState.lastSeenVersion;
    let message = "No change since the last check.";

    if (fileChanged) {
        const outcome = checkSavedFile();
        const farmsTouched = addToInbox(outcome.ready);
        demoState.heldRows = outcome.held;
        demoState.lastSeenVersion = demoState.fileVersion;
        message = `${samplesWord(outcome.ready.length).replace("sample", "new sample")} added to the Sample Inbox, ${rowsWord(outcome.held.length)} ${outcome.held.length === 1 ? "needs" : "need"} fixing.`;
        if (outcome.ready.length === 0 && outcome.held.length === 0) message = "The file changed, but there were no new rows.";
        const farmList = [...farmsTouched.entries()].map(([farmName, count]) => `${farmName} (${samplesWord(count)})`);
        demoState.history.unshift({
            when: checkedAt,
            who: checkKind === "automatic" ? "Automatic check" : "Check now",
            text: farmList.length > 0 ? `${message} ${farmList.join(", ")}.` : message,
        });
    }

    demoState.latestCheck = { checkedAt, message, changedAt: demoState.fileSavedAt };
    demoState.nextCheckAt = addMinutes(checkedAt, CHECK_EVERY_MINUTES);
    renderAll();
}

function saveSheet() {
    demoState.savedRows = copyRows(demoState.sheetRows);
    demoState.fileVersion += 1;
    demoState.fileSavedAt = currentTime();
    demoState.unsavedChanges = false;
    renderAll();
}

function nextBfiCode() {
    const usedNumbers = demoState.sheetRows
        .map((sheetRow) => Number((sheetRow.cells.bfiCode.match(/(\d+)\s*$/) || [])[1]))
        .filter((number) => Number.isFinite(number));
    const highest = Math.max(9100, ...usedNumbers);
    return `BFI-${highest + 1}`;
}

function addRow(example) {
    const bfiCode = nextBfiCode();
    const filled = example || {};
    const cellValues = [
        bfiCode,
        filled.tubeLabel ? `${filled.tubeLabel} ${bfiCode.slice(4)}` : "",
        filled.receivedDate || shortDate(currentTime()),
        filled.category || "",
        filled.species || "",
        filled.cut || "",
        filled.breed || "",
        filled.practice || "",
        filled.farmName || "",
    ];
    demoState.sheetRows.push(makeSheetRow(cellValues));
    demoState.unsavedChanges = true;
    renderAll();
    const tableWrap = document.querySelector("#sheet-table-wrap");
    if (tableWrap) tableWrap.scrollTop = tableWrap.scrollHeight;
}

function addGoodExample() {
    const example = DEMO_GOOD_EXAMPLES[demoState.goodExampleIndex % DEMO_GOOD_EXAMPLES.length];
    demoState.goodExampleIndex += 1;
    addRow(example);
}

function addProblemExample() {
    const example = DEMO_PROBLEM_EXAMPLES[demoState.problemExampleIndex % DEMO_PROBLEM_EXAMPLES.length];
    demoState.problemExampleIndex += 1;
    addRow(example);
}

function removeRow(rowId) {
    demoState.sheetRows = demoState.sheetRows.filter((sheetRow) => sheetRow.rowId !== rowId);
    demoState.unsavedChanges = true;
    renderAll();
}

function skipAheadOneHour() {
    const minutesToNextCheck = Math.max(1, Math.ceil((demoState.nextCheckAt - currentTime()) / 60000));
    demoState.skippedMinutes += minutesToNextCheck;
    runCheck("automatic");
}

function approveCard(cardId) {
    const farmCard = demoState.farmCards.find((knownCard) => knownCard.cardId === cardId);
    const cardSamples = demoState.waitingSamples.filter((sample) => sample.cardId === cardId);
    const projectName = `${farmCard.farmName} samples ${shortDate(currentTime())}`;
    cardSamples.forEach((sample) => {
        demoState.liveSamples.push({ ...sample, projectName, approvedAt: currentTime() });
    });
    demoState.waitingSamples = demoState.waitingSamples.filter((sample) => sample.cardId !== cardId);
    farmCard.status = "approved";
    demoState.history.unshift({
        when: currentTime(),
        who: "Approve",
        text: `${samplesWord(cardSamples.length)} from ${farmCard.farmName} approved. Project "${projectName}" created.`,
    });
    closeCard();
    renderAll();
    document.querySelector("#live-samples").scrollIntoView({ behavior: "smooth", block: "start" });
}

function rejectCard(cardId) {
    const farmCard = demoState.farmCards.find((knownCard) => knownCard.cardId === cardId);
    const cardSamples = demoState.waitingSamples.filter((sample) => sample.cardId === cardId);
    const confirmed = window.confirm(
        `Reject ${samplesWord(cardSamples.length)} from ${farmCard.farmName}? They are removed from the Sample Inbox and nothing goes live.`,
    );
    if (!confirmed) return;
    demoState.waitingSamples = demoState.waitingSamples.filter((sample) => sample.cardId !== cardId);
    farmCard.status = "rejected";
    demoState.history.unshift({
        when: currentTime(),
        who: "Reject",
        text: `${samplesWord(cardSamples.length)} from ${farmCard.farmName} rejected.`,
    });
    closeCard();
    renderAll();
}

function rowStatus(sheetRow, savedIndex) {
    const key = codeKey(sheetRow.cells.bfiCode);
    if (demoState.existingCodes.has(key)) return { label: "Already existed", tone: "grey" };
    const savedRow = demoState.savedRows[savedIndex];
    const savedSame = savedRow && savedRow.rowId === sheetRow.rowId && JSON.stringify(savedRow.cells) === JSON.stringify(sheetRow.cells);
    if (!savedSame) return { label: "Not saved yet", tone: "blue" };
    if (demoState.liveSamples.some((sample) => sample.bfiCode === key)) return { label: "Live", tone: "green" };
    if (demoState.waitingSamples.some((sample) => sample.bfiCode === key)) return { label: "Waiting for approval", tone: "purple" };
    if (demoState.heldRows.some((heldRow) => heldRow.sheetRowNumber === savedIndex + 2)) return { label: "Needs fixing", tone: "amber" };
    if (demoState.fileVersion !== demoState.lastSeenVersion) return { label: "Saved, not checked yet", tone: "blue" };
    return { label: "Not picked up", tone: "grey" };
}

function renderSheet() {
    const headerCells = SHEET_COLUMNS.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("");
    const bodyRows = demoState.sheetRows
        .map((sheetRow, rowIndex) => {
            const status = rowStatus(sheetRow, rowIndex);
            const isExisting = status.tone === "grey" && status.label === "Already existed";
            const inputCells = SHEET_COLUMNS.map(
                (column) =>
                    `<td><input data-row-id="${sheetRow.rowId}" data-cell-key="${column.key}" value="${escapeHtml(sheetRow.cells[column.key])}" aria-label="Row ${rowIndex + 2} ${escapeHtml(column.label)}"></td>`,
            ).join("");
            const removeButton = isExisting
                ? ""
                : `<button type="button" class="link-button" data-remove-row="${sheetRow.rowId}" aria-label="Delete row ${rowIndex + 2}">Delete</button>`;
            return `<tr class="${isExisting ? "existing-row" : ""}">
                <td class="row-number">${rowIndex + 2}</td>
                <td class="side-cell"><span class="pill pill-${status.tone}">${status.label}</span></td>
                ${inputCells}
                <td class="side-cell">${removeButton}</td>
            </tr>`;
        })
        .join("");

    document.querySelector("#sheet-table-wrap").innerHTML = `
        <table class="sheet-table">
            <thead><tr><th class="row-number">1</th><th class="status-head">What the website did</th>${headerCells}<th></th></tr></thead>
            <tbody>${bodyRows}</tbody>
        </table>`;

    const saveState = document.querySelector("#save-state");
    if (demoState.unsavedChanges) {
        saveState.textContent = "You have changes that are not saved. The website only sees the saved file.";
        saveState.className = "save-state unsaved";
    } else {
        saveState.textContent = demoState.fileSavedAt
            ? `Saved to Box at ${shortDateTime(demoState.fileSavedAt)}.`
            : "Saved. These are the rows the lab already had.";
        saveState.className = "save-state";
    }
    document.querySelector("#save-button").disabled = !demoState.unsavedChanges;
}

function renderCheckStatus() {
    const latestCheck = demoState.latestCheck;
    const changedLine = latestCheck.changedAt
        ? `<div class="small muted">File last changed: ${shortDateTime(latestCheck.changedAt)}</div>`
        : "";
    document.querySelector("#check-status").innerHTML = `
        <div><span class="muted">Last checked: </span>${shortDateTime(latestCheck.checkedAt)} <span class="muted">(Box: Sample_Master_File.xlsx)</span></div>
        <div>${escapeHtml(latestCheck.message)}</div>
        ${changedLine}
        <div class="small muted">Next automatic check: ${shortDateTime(demoState.nextCheckAt)}</div>`;
    document.querySelector("#demo-clock").textContent = shortDateTime(currentTime());
}

function renderHeldRows() {
    const heldBox = document.querySelector("#held-rows");
    if (demoState.heldRows.length === 0) {
        heldBox.innerHTML = "";
        return;
    }
    const heldLines = demoState.heldRows
        .map(
            (heldRow) => `<tr>
                <td>${heldRow.sheetRowNumber}</td>
                <td class="mono">${escapeHtml(heldRow.bfiCode || "none")}</td>
                <td>${heldRow.reasons.map((reason) => `<div>${escapeHtml(reason)}</div>`).join("")}</td>
            </tr>`,
        )
        .join("");
    heldBox.innerHTML = `
        <div class="held-box">
            <p class="strong">Rows that need fixing in the Excel file. Once fixed, the next check picks them up.</p>
            <div class="table-scroll">
                <table class="held-table">
                    <thead><tr><th>Row</th><th>BFI code</th><th>Reason</th></tr></thead>
                    <tbody>${heldLines}</tbody>
                </table>
            </div>
        </div>`;
}

function renderInbox() {
    const waitingCards = demoState.farmCards.filter((farmCard) => farmCard.status === "waiting");
    const waitingTotal = demoState.waitingSamples.length;
    document.querySelector("#inbox-count").textContent =
        waitingTotal > 0 ? `${samplesWord(waitingTotal)} waiting for approval` : "";
    const inboxCards = document.querySelector("#inbox-cards");
    if (waitingCards.length === 0) {
        inboxCards.innerHTML = `<p class="muted small">No samples are waiting for approval.</p>`;
        return;
    }
    inboxCards.innerHTML = waitingCards
        .map((farmCard) => {
            const cardSamples = demoState.waitingSamples.filter((sample) => sample.cardId === farmCard.cardId);
            const receivedDates = cardSamples.map((sample) => sample.receivedDate.getTime());
            const firstDate = shortDate(new Date(Math.min(...receivedDates)));
            const lastDate = shortDate(new Date(Math.max(...receivedDates)));
            const receivedText = firstDate === lastDate ? firstDate : `${firstDate} to ${lastDate}`;
            return `<button type="button" class="farm-card" data-open-card="${farmCard.cardId}">
                <span class="farm-card-name">${escapeHtml(farmCard.farmName)}</span>
                <span class="small">${samplesWord(cardSamples.length)} waiting</span>
                <span class="small muted">Received ${receivedText}</span>
                <span class="farm-card-action">Review and approve</span>
            </button>`;
        })
        .join("");
}

function renderHistory() {
    const historyList = document.querySelector("#history-list");
    if (demoState.history.length === 0) {
        historyList.innerHTML = `<li class="muted">Nothing yet.</li>`;
        return;
    }
    historyList.innerHTML = demoState.history
        .map(
            (historyLine) =>
                `<li><span class="muted">${shortDateTime(historyLine.when)}</span> <span class="strong">${escapeHtml(historyLine.who)}:</span> ${escapeHtml(historyLine.text)}</li>`,
        )
        .join("");
}

function renderLiveSamples() {
    const liveBox = document.querySelector("#live-table");
    if (demoState.liveSamples.length === 0) {
        liveBox.innerHTML = `<p class="muted small">No samples have been approved in this demo yet.</p>`;
        return;
    }
    const liveLines = demoState.liveSamples
        .map(
            (sample) => `<tr>
                <td class="mono">${escapeHtml(sample.bfiCode)}</td>
                <td>${escapeHtml(sample.tubeLabel)}</td>
                <td>${escapeHtml(sample.farmName)}</td>
                <td>${escapeHtml(sample.foodDescription)}</td>
                <td>${escapeHtml(sample.projectName)}</td>
                <td>${shortDateTime(sample.approvedAt)}</td>
            </tr>`,
        )
        .join("");
    liveBox.innerHTML = `
        <div class="table-scroll">
            <table class="plain-table">
                <thead><tr><th>BFI code</th><th>Tube label</th><th>Farm</th><th>Food</th><th>Project</th><th>Approved</th></tr></thead>
                <tbody>${liveLines}</tbody>
            </table>
        </div>`;
}

function renderAll() {
    renderSheet();
    renderCheckStatus();
    renderHeldRows();
    renderInbox();
    renderHistory();
    renderLiveSamples();
}

function openCard(cardId) {
    const farmCard = demoState.farmCards.find((knownCard) => knownCard.cardId === cardId);
    const cardSamples = demoState.waitingSamples.filter((sample) => sample.cardId === cardId);
    demoState.openCardId = cardId;
    const sampleLines = cardSamples
        .map(
            (sample) => `<tr>
                <td class="mono">${escapeHtml(sample.bfiCode)}</td>
                <td>${escapeHtml(sample.tubeLabel)}</td>
                <td>${escapeHtml(sample.sampleType)}</td>
                <td>${escapeHtml(sample.foodDescription)}</td>
                <td>${escapeHtml(sample.practices.join(", "))}</td>
                <td>${shortDate(sample.receivedDate)}</td>
            </tr>`,
        )
        .join("");
    const dialog = document.querySelector("#farm-card-dialog");
    dialog.innerHTML = `
        <div class="dialog-head">
            <div>
                <h3>${escapeHtml(farmCard.farmName)}</h3>
                <p class="muted small">Check the samples below. Nothing is live until you press Approve.</p>
            </div>
            <button type="button" class="close-button" data-close-card aria-label="Close">&times;</button>
        </div>
        <div class="table-scroll">
            <table class="plain-table">
                <thead><tr><th>BFI code</th><th>Tube label</th><th>Type</th><th>Food</th><th>Practices</th><th>Received</th></tr></thead>
                <tbody>${sampleLines}</tbody>
            </table>
        </div>
        <div class="dialog-actions">
            <button type="button" class="outline-button" data-reject-card="${cardId}">Reject</button>
            <button type="button" class="dark-button" data-approve-card="${cardId}">Approve ${samplesWord(cardSamples.length)}</button>
        </div>`;
    dialog.showModal();
}

function closeCard() {
    const dialog = document.querySelector("#farm-card-dialog");
    if (dialog.open) dialog.close();
    demoState.openCardId = null;
}

function handleSheetTyping(inputEvent) {
    const input = inputEvent.target;
    if (!input.dataset.rowId) return;
    const sheetRow = demoState.sheetRows.find((knownRow) => knownRow.rowId === Number(input.dataset.rowId));
    sheetRow.cells[input.dataset.cellKey] = input.value;
    if (!demoState.unsavedChanges) {
        demoState.unsavedChanges = true;
        const saveState = document.querySelector("#save-state");
        saveState.textContent = "You have changes that are not saved. The website only sees the saved file.";
        saveState.className = "save-state unsaved";
        document.querySelector("#save-button").disabled = false;
    }
}

function handleClicks(clickEvent) {
    const clicked = clickEvent.target.closest("button");
    if (!clicked) return;
    const { removeRow: removeRowId, openCard: openCardId, approveCard: approveCardId, rejectCard: rejectCardId } = clicked.dataset;
    if (removeRowId) removeRow(Number(removeRowId));
    else if (openCardId) openCard(Number(openCardId));
    else if (approveCardId) approveCard(Number(approveCardId));
    else if (rejectCardId) rejectCard(Number(rejectCardId));
    else if ("closeCard" in clicked.dataset) closeCard();
}

function startDemo() {
    demoState.sheetRows = DEMO_EXISTING_ROWS.map((cellValues) => makeSheetRow(cellValues));
    demoState.savedRows = copyRows(demoState.sheetRows);
    demoState.sheetRows.forEach((sheetRow) => demoState.existingCodes.add(codeKey(sheetRow.cells.bfiCode)));

    const firstReadAt = addMinutes(currentTime(), -45);
    demoState.history.push({
        when: firstReadAt,
        who: "First read",
        text: `First read of the file: ${demoState.existingCodes.size} BFI codes marked as existing. New rows added from now on will be picked up.`,
    });
    demoState.latestCheck = { checkedAt: addMinutes(currentTime(), -15), message: "No change since the last check.", changedAt: null };
    demoState.nextCheckAt = addMinutes(demoState.latestCheck.checkedAt, CHECK_EVERY_MINUTES);

    document.querySelector("#add-good-row").addEventListener("click", addGoodExample);
    document.querySelector("#add-problem-row").addEventListener("click", addProblemExample);
    document.querySelector("#add-empty-row").addEventListener("click", () => addRow(null));
    document.querySelector("#save-button").addEventListener("click", saveSheet);
    document.querySelector("#check-now").addEventListener("click", () => runCheck("button"));
    document.querySelector("#skip-hour").addEventListener("click", skipAheadOneHour);
    document.querySelector("#sheet-table-wrap").addEventListener("input", handleSheetTyping);
    document.addEventListener("click", handleClicks);
    document.querySelector("#farm-card-dialog").addEventListener("close", () => {
        demoState.openCardId = null;
    });
    renderAll();
}

startDemo();
