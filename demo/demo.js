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
    { key: "storage", label: "Storage Location Master", defaultValue: "-80 Freezer A" },
];

const CHECK_EVERY_MINUTES = 60;
const LAB_TIME_ZONE = "America/Denver";
const SOURCE_DESCRIPTION = "Box file 1350482221";
const FILE_NAME = "Sample_Master_File.xlsx";

const BUTTON_BASE =
    "inline-flex items-center justify-center gap-2 whitespace-nowrap rounded-md text-sm font-medium ring-offset-background transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:pointer-events-none disabled:opacity-50 [&_svg]:pointer-events-none [&_svg]:size-4 [&_svg]:shrink-0 h-10 px-4 py-2";
const BUTTON_PRIMARY = `${BUTTON_BASE} bg-primary text-primary-foreground hover:bg-primary/90`;
const BUTTON_OUTLINE = `${BUTTON_BASE} border border-input bg-background hover:bg-accent hover:text-accent-foreground`;
const BUTTON_DESTRUCTIVE = `${BUTTON_BASE} bg-destructive text-destructive-foreground hover:bg-destructive/90`;

const ICON_START = `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"`;
const ICONS = {
    refresh: `${ICON_START} class="lucide lucide-refresh-cw mr-2 h-4 w-4"><path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"></path><path d="M21 3v5h-5"></path><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"></path><path d="M8 16H3v5"></path></svg>`,
    loader: `${ICON_START} class="lucide lucide-loader-circle mr-2 h-4 w-4 animate-spin"><path d="M21 12a9 9 0 1 1-6.219-8.56"></path></svg>`,
    chevronRight: `${ICON_START} class="lucide lucide-chevron-right mr-1 h-4 w-4"><path d="m9 18 6-6-6-6"></path></svg>`,
    chevronDown: `${ICON_START} class="lucide lucide-chevron-down mr-1 h-4 w-4"><path d="m6 9 6 6 6-6"></path></svg>`,
    upload: `${ICON_START} class="lucide lucide-upload mr-2 h-4 w-4"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" x2="12" y1="3" y2="15"></line></svg>`,
    close: `${ICON_START} class="lucide lucide-x h-4 w-4"><path d="M18 6 6 18"></path><path d="m6 6 12 12"></path></svg>`,
};

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
    rejectedCodes: new Set(),
    latestCheck: null,
    nextCheckAt: null,
    recentUploads: [],
    nextRowId: 1,
    nextCardId: 1,
    nextGroupId: 1201,
    goodExampleIndex: 0,
    problemExampleIndex: 0,
    checking: false,
    showManualUpload: false,
    showHistory: false,
    openCard: null,
    toastTimer: null,
};

function currentTime() {
    return addMinutes(new Date(), demoState.skippedMinutes);
}

function addMinutes(when, minutes) {
    return new Date(when.getTime() + minutes * 60000);
}

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

function shortDate(when) {
    return when.toLocaleDateString("en-US", { timeZone: LAB_TIME_ZONE });
}

function shortDateTime(when) {
    return when.toLocaleString("en-US", {
        timeZone: LAB_TIME_ZONE,
        month: "numeric",
        day: "numeric",
        year: "numeric",
        hour: "numeric",
        minute: "2-digit",
    });
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
    const parsed = new Date(year, month - 1, day, 12);
    if (parsed.getMonth() !== month - 1 || parsed.getDate() !== day) return null;
    return parsed;
}

function makeSheetRow(cellValues) {
    const cells = {};
    SHEET_COLUMNS.forEach((column, columnIndex) => {
        cells[column.key] = cellValues[columnIndex] ?? column.defaultValue ?? "";
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
            sampleType,
            foodDescription: `${food.foodName}, ${breed}, ${cut}`,
            foodName: food.foodName,
            cutName: cut,
            practices,
            receivedDate,
            storage: cells.storage.trim(),
        },
    };
}

function looksLikeSample(cells) {
    return ["bfiCode", "tubeLabel", "farmName", "category", "species"].some((cellKey) => cells[cellKey].trim());
}

function checkSavedFile() {
    const liveCodes = new Set(demoState.liveSamples.map((sample) => sample.bfiCode));
    const waitingCodes = new Set(demoState.waitingSamples.map((sample) => sample.bfiCode));
    const sampleRows = demoState.savedRows
        .map((sheetRow, rowIndex) => ({ sheetRowNumber: rowIndex + 4, cells: sheetRow.cells }))
        .filter((sheetRow) => looksLikeSample(sheetRow.cells));

    const rowsByCode = new Map();
    sampleRows.forEach((sheetRow) => {
        const key = codeKey(sheetRow.cells.bfiCode);
        if (!key) return;
        if (!rowsByCode.has(key)) rowsByCode.set(key, []);
        rowsByCode.get(key).push(sheetRow.sheetRowNumber);
    });

    const outcome = { oldRows: 0, alreadyInSystem: 0, ready: [], held: [] };
    sampleRows.forEach((sheetRow) => {
        const key = codeKey(sheetRow.cells.bfiCode);
        const heldRow = (reasons) => ({ sheetRowNumber: sheetRow.sheetRowNumber, bfiCode: key, reasons });
        if (!key) {
            outcome.held.push(heldRow(["BFI_Code is empty."]));
        } else if (demoState.existingCodes.has(key)) {
            outcome.oldRows += 1;
        } else if (liveCodes.has(key) || waitingCodes.has(key)) {
            outcome.alreadyInSystem += 1;
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
    const newCard = { cardId: demoState.nextCardId, farmName, status: "waiting" };
    demoState.nextCardId += 1;
    demoState.farmCards.push(newCard);
    return newCard;
}

function addToInbox(readySamples) {
    readySamples.forEach((sample) => {
        const farmCard = pendingCardForFarm(sample.farmName);
        demoState.rejectedCodes.delete(sample.bfiCode);
        demoState.waitingSamples.push({ ...sample, cardId: farmCard.cardId });
    });
}

function runCheck() {
    const checkedAt = currentTime();
    const fileChanged = demoState.fileVersion !== demoState.lastSeenVersion;
    let message = "No change since the last check.";

    if (fileChanged) {
        const outcome = checkSavedFile();
        addToInbox(outcome.ready);
        demoState.heldRows = outcome.held;
        demoState.lastSeenVersion = demoState.fileVersion;
        const heldVerb = outcome.held.length === 1 ? "needs" : "need";
        message = `${samplesWord(outcome.ready.length).replace("sample", "new sample")} added to the Sample Inbox, ${rowsWord(outcome.held.length)} ${heldVerb} fixing.`;
        demoState.recentUploads.unshift({
            uploadedAt: checkedAt,
            uploadedBy: "",
            summary: `${outcome.ready.length} new, ${outcome.held.length} held back, ${outcome.alreadyInSystem} already in the system`,
        });
    }

    demoState.latestCheck = { checkedAt, message, changedAt: demoState.fileSavedAt };
    demoState.nextCheckAt = addMinutes(checkedAt, CHECK_EVERY_MINUTES);
    renderAll();
}

function pressCheckNow() {
    if (demoState.checking) return;
    demoState.checking = true;
    renderStatusArea();
    window.setTimeout(() => {
        demoState.checking = false;
        runCheck();
    }, 700);
}

function skipAheadOneHour() {
    const minutesToNextCheck = Math.max(1, Math.ceil((demoState.nextCheckAt - currentTime()) / 60000));
    demoState.skippedMinutes += minutesToNextCheck;
    runCheck();
    showToast(`It is now ${shortDateTime(currentTime())}. The hourly check ran by itself.`);
}

function saveSheet() {
    demoState.savedRows = copyRows(demoState.sheetRows);
    demoState.fileVersion += 1;
    demoState.fileSavedAt = currentTime();
    demoState.unsavedChanges = false;
    renderSheet();
    showToast("Saved to Box. Close the file and press Check now, or wait for the hourly check.");
}

function nextBfiCode() {
    const usedNumbers = demoState.sheetRows
        .map((sheetRow) => Number((sheetRow.cells.bfiCode.match(/(\d+)\s*$/) || [])[1]))
        .filter((number) => Number.isFinite(number));
    return `BFI-${Math.max(9100, ...usedNumbers) + 1}`;
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
    renderSheet();
    const tableWrap = document.querySelector("#sheet-table-wrap");
    tableWrap.scrollTop = tableWrap.scrollHeight;
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
    renderSheet();
}

function cardSamples(cardId) {
    return demoState.waitingSamples.filter((sample) => sample.cardId === cardId);
}

function approveCard(cardId) {
    const farmCard = demoState.farmCards.find((knownCard) => knownCard.cardId === cardId);
    const approvedSamples = cardSamples(cardId);
    const projectName = `${farmCard.farmName} samples ${shortDate(currentTime())}`;
    approvedSamples.forEach((sample) => demoState.liveSamples.push({ ...sample, projectName }));
    demoState.waitingSamples = demoState.waitingSamples.filter((sample) => sample.cardId !== cardId);
    farmCard.status = "approved";
    const groups = buildGroups(approvedSamples);
    demoState.openCard = {
        cardId,
        farmName: farmCard.farmName,
        approvedSamples,
        projectName,
        confirmingReject: false,
        groupScreen: { groups, savedGroups: copyGroups(groups), editing: false },
    };
    renderAll();
}

function buildGroups(samples) {
    const groupsByKey = new Map();
    samples.forEach((sample) => {
        const groupKey = `${sample.foodDescription}|${[...sample.practices].sort().join(",")}`;
        if (!groupsByKey.has(groupKey)) {
            groupsByKey.set(groupKey, { groupId: demoState.nextGroupId, groupName: "Unknown", isBenchmark: false, codes: [] });
            demoState.nextGroupId += 1;
        }
        groupsByKey.get(groupKey).codes.push(sample.bfiCode);
    });
    return [...groupsByKey.values()];
}

function copyGroups(groups) {
    return groups.map((group) => ({ ...group, codes: [...group.codes] }));
}

function rejectCard(cardId) {
    const farmCard = demoState.farmCards.find((knownCard) => knownCard.cardId === cardId);
    cardSamples(cardId).forEach((sample) => demoState.rejectedCodes.add(sample.bfiCode));
    demoState.waitingSamples = demoState.waitingSamples.filter((sample) => sample.cardId !== cardId);
    farmCard.status = "rejected";
    demoState.openCard = null;
    renderAll();
}

function rowStatus(sheetRow, rowIndex) {
    const key = codeKey(sheetRow.cells.bfiCode);
    if (demoState.existingCodes.has(key)) return { label: "Already existed", tone: "grey" };
    const savedRow = demoState.savedRows[rowIndex];
    const savedSame = savedRow && savedRow.rowId === sheetRow.rowId && JSON.stringify(savedRow.cells) === JSON.stringify(sheetRow.cells);
    if (!savedSame) return { label: "Not saved yet", tone: "blue" };
    if (demoState.liveSamples.some((sample) => sample.bfiCode === key)) return { label: "Live", tone: "green" };
    if (demoState.waitingSamples.some((sample) => sample.bfiCode === key)) return { label: "In the Sample Inbox", tone: "purple" };
    if (demoState.heldRows.some((heldRow) => heldRow.sheetRowNumber === rowIndex + 4)) return { label: "Needs fixing", tone: "amber" };
    if (demoState.fileVersion !== demoState.lastSeenVersion) return { label: "Saved, not checked yet", tone: "blue" };
    if (demoState.rejectedCodes.has(key)) return { label: "Rejected, comes back after the next save", tone: "amber" };
    return { label: "Not picked up", tone: "grey" };
}

function renderSheet() {
    const headerCells = SHEET_COLUMNS.map((column) => `<th>${escapeHtml(column.label)}</th>`).join("");
    const bodyRows = demoState.sheetRows
        .map((sheetRow, rowIndex) => {
            const sheetRowNumber = rowIndex + 4;
            const status = rowStatus(sheetRow, rowIndex);
            const isExisting = status.label === "Already existed";
            const inputCells = SHEET_COLUMNS.map(
                (column) =>
                    `<td><input data-row-id="${sheetRow.rowId}" data-cell-key="${column.key}" value="${escapeHtml(sheetRow.cells[column.key])}" aria-label="Row ${sheetRowNumber} ${escapeHtml(column.label)}"></td>`,
            ).join("");
            const removeButton = isExisting
                ? ""
                : `<button type="button" class="sheet-delete" data-remove-row="${sheetRow.rowId}" aria-label="Delete row ${sheetRowNumber}">Delete</button>`;
            return `<tr class="${isExisting ? "existing-row" : ""}">
                <td class="row-number">${sheetRowNumber}</td>
                <td class="side-cell"><span class="pill pill-${status.tone}">${status.label}</span></td>
                ${inputCells}
                <td class="side-cell">${removeButton}</td>
            </tr>`;
        })
        .join("");

    document.querySelector("#sheet-table-wrap").innerHTML = `
        <table class="sheet-table">
            <thead><tr><th class="row-number">3</th><th class="status-head">What the website did</th>${headerCells}<th></th></tr></thead>
            <tbody>${bodyRows}</tbody>
        </table>`;
    renderSaveState();
}

function renderSaveState() {
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

function heldRowsMarkup() {
    if (demoState.heldRows.length === 0) return "";
    const heldLines = demoState.heldRows
        .map(
            (heldRow) => `<tr class="border-t border-amber-200 align-top">
                <td class="py-1 pr-3">${heldRow.sheetRowNumber}</td>
                <td class="py-1 pr-3 font-mono">${escapeHtml(heldRow.bfiCode || "none")}</td>
                <td class="py-1">${heldRow.reasons.map((reason) => `<div>${escapeHtml(reason)}</div>`).join("")}</td>
            </tr>`,
        )
        .join("");
    return `<div class="rounded border border-amber-300 bg-amber-50 p-3 text-amber-900">
        <p class="text-sm font-semibold">Rows that need fixing in the Excel file. Once fixed, the next check picks them up.</p>
        <table class="mt-2 w-full text-sm">
            <thead><tr class="text-left text-xs"><th class="py-1 pr-3">Row</th><th class="py-1 pr-3">BFI code</th><th class="py-1">Reason</th></tr></thead>
            <tbody>${heldLines}</tbody>
        </table>
    </div>`;
}

function renderStatusArea() {
    const latestCheck = demoState.latestCheck;
    const changedLine = latestCheck.changedAt
        ? `<div class="text-xs text-muted-foreground">File last changed: ${shortDateTime(latestCheck.changedAt)}</div>`
        : "";
    document.querySelector("#intake-status-area").innerHTML = `<div class="space-y-4">
        <div class="flex flex-wrap items-start justify-between gap-3 rounded border p-4">
            <div class="space-y-1 text-sm">
                <div><span class="text-muted-foreground">Last checked: </span>${shortDateTime(latestCheck.checkedAt)}<span class="text-muted-foreground"> (${SOURCE_DESCRIPTION})</span></div>
                <div>${escapeHtml(latestCheck.message)}</div>
                ${changedLine}
            </div>
            <button type="button" class="${BUTTON_PRIMARY}" id="check-now" ${demoState.checking ? "disabled" : ""}>${demoState.checking ? ICONS.loader : ICONS.refresh}Check now</button>
        </div>
        ${heldRowsMarkup()}
    </div>`;
}

function renderInboxArea() {
    const waitingCards = demoState.farmCards.filter((farmCard) => farmCard.status === "waiting");
    const cardsMarkup =
        waitingCards.length === 0
            ? `<p class="text-sm text-muted-foreground">No samples are waiting for approval.</p>`
            : `<div class="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-3">${waitingCards
                  .map((farmCard) => {
                      const samples = cardSamples(farmCard.cardId);
                      const receivedTimes = samples.map((sample) => sample.receivedDate.getTime());
                      const firstDate = shortDate(new Date(Math.min(...receivedTimes)));
                      const lastDate = shortDate(new Date(Math.max(...receivedTimes)));
                      const receivedText = firstDate === lastDate ? `Received ${firstDate}` : `Received ${firstDate} to ${lastDate}`;
                      return `<button type="button" data-open-card="${farmCard.cardId}" class="rounded-lg border bg-background p-4 text-left shadow-sm transition hover:border-primary hover:shadow">
                          <div class="font-semibold">${escapeHtml(farmCard.farmName)}</div>
                          <div class="mt-1 text-sm">${samplesWord(samples.length)} waiting</div>
                          <div class="mt-1 text-xs text-muted-foreground">${receivedText}</div>
                          <div class="mt-3 text-sm font-medium text-primary">Review and approve</div>
                      </button>`;
                  })
                  .join("")}</div>`;
    document.querySelector("#intake-inbox-area").innerHTML = `
        <h3 class="mb-3 text-lg font-semibold">Sample Inbox <span class="text-sm font-normal text-muted-foreground">${samplesWord(demoState.waitingSamples.length)} waiting for approval</span></h3>
        ${cardsMarkup}`;
}

function renderUploadArea() {
    const uploadBody = demoState.showManualUpload
        ? `<div class="mt-3"><div class="space-y-4">
            <p class="text-sm text-muted-foreground">Only needed if the automatic check cannot reach the file. Choose the whole Sample Master File (.xlsx or .csv); only new rows are added.</p>
            <div class="flex flex-wrap items-center gap-3">
                <input type="file" accept=".xlsx,.xlsm,.csv" class="text-sm">
                <button type="button" class="${BUTTON_OUTLINE}" data-demo-only="In this demo the file always comes from Box. Use the Excel file at the bottom of the screen.">${ICONS.upload}Upload</button>
            </div>
        </div></div>`
        : "";
    document.querySelector("#intake-upload-area").innerHTML = `
        <button type="button" class="flex items-center text-sm font-medium" data-toggle="upload">${demoState.showManualUpload ? ICONS.chevronDown : ICONS.chevronRight}Upload a file by hand</button>
        ${uploadBody}`;
}

function renderHistoryArea() {
    const historyArea = document.querySelector("#intake-history-area");
    const historyLines = demoState.recentUploads
        .map(
            (upload) => `<tr class="border-b">
                <td class="py-1 pr-3 whitespace-nowrap">${shortDateTime(upload.uploadedAt)}</td>
                <td class="py-1 pr-3">${escapeHtml(upload.uploadedBy || "Automatic check")}</td>
                <td class="py-1 pr-3">${FILE_NAME}</td>
                <td class="py-1">${escapeHtml(upload.summary)}</td>
            </tr>`,
        )
        .join("");
    historyArea.innerHTML = `
        <button type="button" class="flex items-center text-sm font-medium" data-toggle="history">${demoState.showHistory ? ICONS.chevronDown : ICONS.chevronRight}Recent checks and uploads</button>
        ${demoState.showHistory ? `<table class="mt-2 w-full text-sm"><tbody>${historyLines}</tbody></table>` : ""}`;
}

function samplesTableMarkup(samples) {
    const sampleLines = samples
        .map(
            (sample) => `<tr class="border-b align-top">
                <td class="py-2 pr-3 font-mono">${escapeHtml(sample.bfiCode)}</td>
                <td class="py-2 pr-3">${escapeHtml(sample.tubeLabel)}</td>
                <td class="py-2 pr-3">${escapeHtml(sample.sampleType)}</td>
                <td class="py-2 pr-3">${escapeHtml(sample.foodDescription)}</td>
                <td class="py-2 pr-3">${escapeHtml(sample.practices.join(", "))}</td>
                <td class="py-2 pr-3">${shortDate(sample.receivedDate)}</td>
                <td class="py-2">${escapeHtml(sample.storage)}</td>
            </tr>`,
        )
        .join("");
    return `<div class="overflow-x-auto">
        <table class="w-full text-sm">
            <thead>
                <tr class="border-b text-left text-xs text-muted-foreground">
                    <th class="py-2 pr-3">BFI code</th><th class="py-2 pr-3">Tube label</th><th class="py-2 pr-3">Type</th>
                    <th class="py-2 pr-3">Food</th><th class="py-2 pr-3">Practices</th><th class="py-2 pr-3">Received</th><th class="py-2">Storage</th>
                </tr>
            </thead>
            <tbody>${sampleLines}</tbody>
        </table>
    </div>`;
}

function reviewButtonsMarkup(openCard, samples) {
    if (openCard.confirmingReject) {
        return `<span class="text-sm text-muted-foreground">Remove these samples from the waiting area? They come back on the next upload.</span>
            <button type="button" class="${BUTTON_OUTLINE}" data-keep-card>Keep</button>
            <button type="button" class="${BUTTON_DESTRUCTIVE}" data-reject-card="${openCard.cardId}">Yes, reject</button>`;
    }
    return `<button type="button" class="${BUTTON_OUTLINE}" data-confirm-reject>Reject</button>
        <button type="button" class="${BUTTON_PRIMARY}" data-approve-card="${openCard.cardId}">Approve ${samplesWord(samples.length)}</button>`;
}

function groupLabel(group) {
    return group.groupName || `Group ${group.groupId}`;
}

function isoDate(when) {
    return when.toLocaleDateString("en-CA", { timeZone: LAB_TIME_ZONE });
}

function groupSampleMarkup(sample, groupId, editing) {
    return `<div class="border rounded-md p-2 bg-gray-50" draggable="${editing}" data-sample-code="${escapeHtml(sample.bfiCode)}" data-from-group="${groupId}">
        <div class="text-xs text-gray-500 mb-1">${escapeHtml(sample.bfiCode)}</div>
        <div class="text-sm font-medium truncate" title="${escapeHtml(sample.tubeLabel)}">${escapeHtml(sample.tubeLabel)}</div>
        <div class="mt-1 grid grid-cols-2 gap-x-2 gap-y-1 text-xs">
            <div class="text-gray-500">Date</div><div>${isoDate(sample.receivedDate)}</div>
            <div class="text-gray-500">Species</div><div>${escapeHtml(sample.foodName)}</div>
            <div class="text-gray-500">Cut</div><div>${escapeHtml(sample.cutName)}</div>
            <div class="text-gray-500">Practice</div><div>${escapeHtml(sample.practices.join(", "))}</div>
        </div>
    </div>`;
}

function groupCardMarkup(group, samplesByCode, editing) {
    const groupName = editing
        ? `<input type="text" class="w-full bg-white border rounded px-2 py-1 text-sm" value="${escapeHtml(group.groupName)}" placeholder="Group ${group.groupId}" data-group-name="${group.groupId}">`
        : `<span class="block truncate" title="${escapeHtml(groupLabel(group))}">${escapeHtml(groupLabel(group))}</span>`;
    const sampleCards = group.codes.length === 0
        ? `<div class="text-sm text-gray-500">No samples in this group.</div>`
        : group.codes.map((bfiCode) => groupSampleMarkup(samplesByCode.get(bfiCode), group.groupId, editing)).join("");
    return `<div class="flex-shrink-0 w-80 border rounded-lg shadow-sm flex flex-col h-full bg-white" data-drop-group="${group.groupId}">
        <div class="px-3 py-2 border-b bg-gray-50 font-semibold">
            ${groupName}
            <label class="mt-2 inline-flex items-center gap-2 text-xs font-normal text-gray-700">
                <input type="checkbox" class="h-4 w-4" data-group-benchmark="${group.groupId}"${group.isBenchmark ? " checked" : ""}${editing ? "" : " disabled"}>
                <span>(Benchmark Group)</span>
            </label>
        </div>
        <div class="p-3 space-y-2 overflow-y-auto flex-1">${sampleCards}</div>
    </div>`;
}

function groupScreenMarkup(openCard) {
    const groupScreen = openCard.groupScreen;
    const samplesByCode = new Map(openCard.approvedSamples.map((sample) => [sample.bfiCode, sample]));
    const editButtons = groupScreen.editing
        ? `<div class="flex flex-col gap-2">
            <button type="button" class="px-3 py-1 rounded border bg-white" data-group-undo>Undo / Cancel</button>
            <button type="button" class="px-3 py-1 rounded border bg-blue-600 text-white" data-group-save>Save</button>
        </div>`
        : `<button type="button" class="px-3 py-1 rounded border bg-blue-50 border-blue-200 text-blue-800 hover:bg-blue-100" data-group-edit title="This will allow you to drag and drop your samples into different groups as well as change the name of any group.">Edit / Move Samples</button>`;
    const groupCards = groupScreen.groups.map((group) => groupCardMarkup(group, samplesByCode, groupScreen.editing)).join("");
    return `<div class="w-full h-full flex flex-col">
        <div class="px-2 pt-1 pb-2 text-xs text-gray-600 flex flex-wrap items-center gap-x-2 gap-y-1">
            <span class="font-medium text-gray-700">Showing this project only</span>
            <span class="text-gray-400">&middot;</span>
            <span>groups and samples belonging to this project</span>
        </div>
        <div class="flex-1 overflow-x-auto">
            <div class="flex gap-4 p-2 h-full items-center">
                <div class="mr-2 flex flex-col gap-2">
                    ${editButtons}
                    <button type="button" class="px-3 py-1 rounded border bg-green-50 border-green-200 text-green-800 hover:bg-green-100" data-group-create title="This will allow you to further delineate your samples by creating and naming a new group that you will be able to move your samples into.">Create New Group</button>
                </div>
                ${groupCards}
            </div>
        </div>
    </div>`;
}

function setGroupEditing(isEditing) {
    demoState.openCard.groupScreen.editing = isEditing;
    renderDialog();
}

function saveGroups() {
    const groupScreen = demoState.openCard.groupScreen;
    groupScreen.savedGroups = copyGroups(groupScreen.groups);
    setGroupEditing(false);
}

function undoGroups() {
    const groupScreen = demoState.openCard.groupScreen;
    groupScreen.groups = copyGroups(groupScreen.savedGroups);
    setGroupEditing(false);
}

function createGroup() {
    const groupScreen = demoState.openCard.groupScreen;
    groupScreen.groups.push({ groupId: demoState.nextGroupId, groupName: "", isBenchmark: false, codes: [] });
    demoState.nextGroupId += 1;
    setGroupEditing(true);
}

function findGroup(groupId) {
    return demoState.openCard.groupScreen.groups.find((group) => group.groupId === Number(groupId));
}

function moveSample(bfiCode, fromGroupId, toGroupId) {
    if (fromGroupId === toGroupId) return;
    const fromGroup = findGroup(fromGroupId);
    const toGroup = findGroup(toGroupId);
    if (!fromGroup || !toGroup || !fromGroup.codes.includes(bfiCode)) return;
    fromGroup.codes = fromGroup.codes.filter((code) => code !== bfiCode);
    toGroup.codes.push(bfiCode);
    renderDialog();
}

function groupScreenIsEditing() {
    return Boolean(demoState.openCard && demoState.openCard.groupScreen && demoState.openCard.groupScreen.editing);
}

function handleDragStart(dragEvent) {
    const sampleCard = dragEvent.target.closest("[data-sample-code]");
    if (!sampleCard || !groupScreenIsEditing()) return;
    dragEvent.dataTransfer.setData("text/plain", JSON.stringify({ bfiCode: sampleCard.dataset.sampleCode, fromGroupId: sampleCard.dataset.fromGroup }));
}

function handleDragOver(dragEvent) {
    if (groupScreenIsEditing() && dragEvent.target.closest("[data-drop-group]")) dragEvent.preventDefault();
}

function handleDrop(dragEvent) {
    const groupCard = dragEvent.target.closest("[data-drop-group]");
    if (!groupCard || !groupScreenIsEditing()) return;
    dragEvent.preventDefault();
    const dragged = JSON.parse(dragEvent.dataTransfer.getData("text/plain") || "{}");
    if (dragged.bfiCode) moveSample(dragged.bfiCode, dragged.fromGroupId, groupCard.dataset.dropGroup);
}

function handleGroupTyping(inputEvent) {
    const field = inputEvent.target;
    if (field.dataset.groupName) findGroup(field.dataset.groupName).groupName = field.value;
    if (field.dataset.groupBenchmark) findGroup(field.dataset.groupBenchmark).isBenchmark = field.checked;
}

function renderDialog() {
    const dialogRoot = document.querySelector("#dialog-root");
    const openCard = demoState.openCard;
    if (!openCard) {
        dialogRoot.innerHTML = "";
        document.body.style.overflow = "";
        return;
    }
    const approved = Boolean(openCard.approvedSamples);
    const samples = approved ? openCard.approvedSamples : cardSamples(openCard.cardId);
    const description = approved
        ? `${samplesWord(samples.length)} approved and added to ${escapeHtml(openCard.projectName)}.`
        : "Check the samples below. Nothing is live until you press Approve.";
    const body = approved
        ? `<div class="space-y-3">
            <p class="text-sm">The samples are live and were grouped automatically. Move them between groups below if needed.</p>
            ${groupScreenMarkup(openCard)}
            <div class="flex justify-end"><button type="button" class="${BUTTON_PRIMARY}" data-close-card>Done</button></div>
        </div>`
        : `${samplesTableMarkup(samples)}
        <div class="flex flex-wrap items-center justify-end gap-2 pt-2">${reviewButtonsMarkup(openCard, samples)}</div>`;

    dialogRoot.innerHTML = `
        <div class="fixed inset-0 z-50 bg-black/80" data-close-card data-state="open"></div>
        <div role="dialog" aria-modal="true" aria-labelledby="farm-card-title" data-state="open" class="fixed left-[50%] top-[50%] z-50 grid w-full translate-x-[-50%] translate-y-[-50%] gap-4 border bg-background p-6 shadow-lg duration-200 sm:rounded-lg max-w-5xl max-h-[90vh] overflow-y-auto">
            <div class="flex flex-col space-y-1.5 text-center sm:text-left">
                <h2 id="farm-card-title" class="text-lg font-semibold leading-none tracking-tight">${escapeHtml(openCard.farmName)}</h2>
                <p class="text-sm text-muted-foreground">${description}</p>
            </div>
            ${body}
            <button type="button" data-close-card class="absolute right-4 top-4 rounded-sm opacity-70 ring-offset-background transition-opacity hover:opacity-100 focus:outline-none focus:ring-2 focus:ring-ring focus:ring-offset-2">${ICONS.close}<span class="sr-only">Close</span></button>
        </div>`;
    document.body.style.overflow = "hidden";
}

function renderAll() {
    renderStatusArea();
    renderInboxArea();
    renderUploadArea();
    renderHistoryArea();
    renderDialog();
    renderSheet();
}

function openFarmCard(cardId) {
    const farmCard = demoState.farmCards.find((knownCard) => knownCard.cardId === cardId);
    demoState.openCard = { cardId, farmName: farmCard.farmName, confirmingReject: false };
    renderDialog();
}

function closeFarmCard() {
    demoState.openCard = null;
    renderDialog();
}

function setExcelOpen(shouldOpen) {
    document.querySelector("#excel-panel").hidden = !shouldOpen;
    document.querySelector("#excel-backdrop").hidden = !shouldOpen;
    if (shouldOpen) renderSheet();
}

function showToast(message) {
    const toast = document.querySelector("#demo-toast");
    toast.textContent = message;
    toast.hidden = false;
    window.clearTimeout(demoState.toastTimer);
    demoState.toastTimer = window.setTimeout(() => {
        toast.hidden = true;
    }, 4000);
}

function handleSheetTyping(inputEvent) {
    const input = inputEvent.target;
    if (!input.dataset.rowId) return;
    const sheetRow = demoState.sheetRows.find((knownRow) => knownRow.rowId === Number(input.dataset.rowId));
    sheetRow.cells[input.dataset.cellKey] = input.value;
    if (!demoState.unsavedChanges) {
        demoState.unsavedChanges = true;
        renderSaveState();
    }
}

function handleClicks(clickEvent) {
    const clicked = clickEvent.target.closest("button, a, [data-close-card]");
    if (!clicked) return;
    const clickedData = clicked.dataset;

    if (clickedData.demoOnly) {
        clickEvent.preventDefault();
        showToast(clickedData.demoOnly);
    } else if (clicked.id === "check-now") pressCheckNow();
    else if (clicked.id === "mobile-menu-button" || clicked.id === "mobile-menu-close") toggleSidebar();
    else if (clickedData.openCard) openFarmCard(Number(clickedData.openCard));
    else if (clickedData.approveCard) approveCard(Number(clickedData.approveCard));
    else if (clickedData.rejectCard) rejectCard(Number(clickedData.rejectCard));
    else if ("confirmReject" in clickedData) setConfirmingReject(true);
    else if ("keepCard" in clickedData) setConfirmingReject(false);
    else if ("closeCard" in clickedData) closeFarmCard();
    else if (clickedData.removeRow) removeRow(Number(clickedData.removeRow));
    else if ("groupEdit" in clickedData) setGroupEditing(true);
    else if ("groupSave" in clickedData) saveGroups();
    else if ("groupUndo" in clickedData) undoGroups();
    else if ("groupCreate" in clickedData) createGroup();
    else if (clickedData.toggle === "upload") {
        demoState.showManualUpload = !demoState.showManualUpload;
        renderUploadArea();
    } else if (clickedData.toggle === "history") {
        demoState.showHistory = !demoState.showHistory;
        renderHistoryArea();
    }
}

function setConfirmingReject(isConfirming) {
    demoState.openCard.confirmingReject = isConfirming;
    renderDialog();
}

function toggleSidebar() {
    const sidebar = document.querySelector("#site-sidebar");
    sidebar.style.transform = sidebar.style.transform ? "" : "translateX(0)";
}

function handleKeys(keyEvent) {
    if (keyEvent.key !== "Escape") return;
    if (demoState.openCard) closeFarmCard();
    else setExcelOpen(false);
}

function startDemo() {
    demoState.sheetRows = DEMO_EXISTING_ROWS.map((cellValues) => makeSheetRow(cellValues));
    demoState.savedRows = copyRows(demoState.sheetRows);
    demoState.sheetRows.forEach((sheetRow) => demoState.existingCodes.add(codeKey(sheetRow.cells.bfiCode)));

    const now = currentTime();
    demoState.recentUploads.push({
        uploadedAt: addMinutes(now, -45),
        uploadedBy: "",
        summary: `First read: ${demoState.existingCodes.size} codes marked as existing`,
    });
    demoState.latestCheck = { checkedAt: addMinutes(now, -15), message: "No change since the last check.", changedAt: null };
    demoState.nextCheckAt = addMinutes(demoState.latestCheck.checkedAt, CHECK_EVERY_MINUTES);

    document.querySelector("#open-excel").addEventListener("click", () => setExcelOpen(true));
    document.querySelector("#close-excel").addEventListener("click", () => setExcelOpen(false));
    document.querySelector("#excel-backdrop").addEventListener("click", () => setExcelOpen(false));
    document.querySelector("#add-good-row").addEventListener("click", addGoodExample);
    document.querySelector("#add-problem-row").addEventListener("click", addProblemExample);
    document.querySelector("#add-empty-row").addEventListener("click", () => addRow(null));
    document.querySelector("#save-button").addEventListener("click", saveSheet);
    document.querySelector("#skip-hour").addEventListener("click", skipAheadOneHour);
    document.querySelector("#sheet-table-wrap").addEventListener("input", handleSheetTyping);
    document.addEventListener("click", handleClicks);
    document.addEventListener("keydown", handleKeys);
    const dialogRoot = document.querySelector("#dialog-root");
    dialogRoot.addEventListener("dragstart", handleDragStart);
    dialogRoot.addEventListener("dragover", handleDragOver);
    dialogRoot.addEventListener("drop", handleDrop);
    dialogRoot.addEventListener("input", handleGroupTyping);
    dialogRoot.addEventListener("change", handleGroupTyping);
    renderAll();
}

startDemo();
