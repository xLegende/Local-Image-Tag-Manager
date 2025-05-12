document.addEventListener('DOMContentLoaded', () => {
    // --- DOM Elements ---
    const loadDatasetBtn = document.getElementById('load-dataset-btn');
    const saveAllBtn = document.getElementById('save-all-btn');
    const batchEditSelectionBtn = document.getElementById('batch-edit-selection-btn');
    const themeToggleBtn = document.getElementById('theme-toggle-btn');
    const statusMessage = document.getElementById('status-message');
    const loadingIndicator = document.getElementById('loading-indicator');
    const imageGrid = document.getElementById('image-grid');
    const searchTagsInput = document.getElementById('search-tags');
    const searchModeSelect = document.getElementById('search-mode');
    const excludeTagsInput = document.getElementById('exclude-tags');
    const filterUntaggedSelect = document.getElementById('filter-untagged');
    const filterBtn = document.getElementById('filter-btn');
    const resetFilterBtn = document.getElementById('reset-filter-btn');
    const renameOldTagInput = document.getElementById('rename-old-tag');
    const renameNewTagInput = document.getElementById('rename-new-tag');
    const renameTagBtn = document.getElementById('rename-tag-btn');
    const removeTagNameInput = document.getElementById('remove-tag-name');
    const removeTagBtn = document.getElementById('remove-tag-btn');
    const addTriggerWordsInput = document.getElementById('add-trigger-words');
    const addTagsBtn = document.getElementById('add-tags-btn');
    const removeDuplicatesBtn = document.getElementById('remove-duplicates-btn');
    
    // Utility DOM Elements
    const findDuplicateFilesBtn = document.getElementById('find-duplicate-files-btn');
    const duplicateFilesOutput = document.getElementById('duplicate-files-output');
    const findOrphanedMissingBtn = document.getElementById('find-orphaned-missing-btn');
    const orphanedMissingOutput = document.getElementById('orphaned-missing-output');
    const findCaseInconsistentTagsBtn = document.getElementById('find-case-inconsistent-tags-btn');
    const caseInconsistentOutput = document.getElementById('case-inconsistent-output');
    const exportUniqueTagsBtn = document.getElementById('export-unique-tags-btn');

    const statsTotalImages = document.getElementById('stat-total-images');
    const statsUniqueTags = document.getElementById('stat-unique-tags');
    const statsModifiedImages = document.getElementById('stat-modified-images');
    const tagFrequencySearchInput = document.getElementById('tag-frequency-search');
    const tagFrequencyList = document.getElementById('tag-frequency-list');
    const displayedCountSpan = document.getElementById('displayed-count');
    const totalCountSpan = document.getElementById('total-count');

    // Single Editor Modal Elements
    const editorModal = document.getElementById('editor-modal');
    const editorModalContent = editorModal.querySelector('.modal-content');
    const editorImagePreview = document.getElementById('editor-image-preview');
    const editorFilename = document.getElementById('editor-filename');
    const editorMetadataDisplay = document.getElementById('editor-metadata-display');
    const editorTagsTextarea = document.getElementById('editor-tags-textarea');
    const editorSaveBtn = document.getElementById('editor-save-btn');
    const editorCancelBtn = document.getElementById('editor-cancel-btn');
    const editorRemoveDupsBtn = document.getElementById('editor-remove-duplicates-btn');
    const editorStatus = document.getElementById('editor-status');
    const editorPrevBtn = document.getElementById('editor-prev-btn');
    const editorNextBtn = document.getElementById('editor-next-btn');
    const editorTagSuggestionsBox = document.getElementById('editor-tag-suggestions');

    // Batch Editor Modal Elements
    const batchEditorModal = document.getElementById('batch-editor-modal');
    const batchEditorSelectedCount = document.getElementById('batch-editor-selected-count');
    const batchEditorCommonTagsDiv = document.getElementById('batch-editor-common-tags');
    const batchEditorPartialTagsDiv = document.getElementById('batch-editor-partial-tags');
    const batchEditorAddTagsInput = document.getElementById('batch-editor-add-tags-input');
    const batchEditorRemoveTagsInput = document.getElementById('batch-editor-remove-tags-input');
    const batchEditorApplyBtn = document.getElementById('batch-editor-apply-btn');
    const batchEditorCancelBtn = document.getElementById('batch-editor-cancel-btn');
    const batchEditorStatus = document.getElementById('batch-editor-status');

    // Sort UI Elements
    const sortPropertySelect = document.getElementById('sort-property');
    const sortOrderSelect = document.getElementById('sort-order');
    const applySortBtn = document.getElementById('apply-sort-btn');

    // Custom Rules UI Elements
    const customRulesListDiv = document.getElementById('custom-rules-list');
    const createNewRuleBtn = document.getElementById('create-new-rule-btn');
    const ruleEditorModal = document.getElementById('rule-editor-modal');
    const ruleEditorTitle = document.getElementById('rule-editor-title');
    const ruleIdInput = document.getElementById('rule-id');
    const ruleNameInput = document.getElementById('rule-name');
    const ruleFindInput = document.getElementById('rule-find');
    const ruleReplaceInput = document.getElementById('rule-replace');
    const ruleIsRegexCheckbox = document.getElementById('rule-is-regex');
    const ruleIsCaseSensitiveCheckbox = document.getElementById('rule-is-case-sensitive');
    const ruleApplyToSelect = document.getElementById('rule-apply-to');
    const ruleSpecificTagsGroup = document.getElementById('rule-specific-tags-group');
    const ruleSpecificTagsInput = document.getElementById('rule-specific-tags');
    const ruleEditorSaveBtn = document.getElementById('rule-editor-save-btn');
    const ruleEditorCancelBtn = document.getElementById('rule-editor-cancel-btn');
    const ruleEditorStatus = document.getElementById('rule-editor-status');
    const bulkApplyCustomRulesSelect = document.getElementById('bulk-apply-custom-rules-select');
    const bulkApplyCustomRulesBtn = document.getElementById('bulk-apply-custom-rules-btn');


    // --- State Variables ---
    let datasetHandle = null;
    let allImageData = []; // Holds { imageHandle, tagHandle, imageName, relativePath, imageUrl, tags, originalTags, modified, id, fileSize, lastModified, imageWidth, imageHeight }
    let allFileSystemEntries = new Map(); // Holds all discovered file/dir handles with their full relative path as key. This is for Orphaned/Missing checks.
    let displayedImageDataIndices = [];
    let currentEditIndex = -1;
    let currentEditOriginalTagsString = '';
    let tagFrequencies = new Map();
    let uniqueTags = new Set();
    let nextItemId = 0;
    let initialTotalCount = 0;
    let debouncedUpdateFrequencyList = null;
    let keyboardFocusIndex = -1;

    let selectedItemIds = new Set();
    let lastInteractedItemId = null;

    let currentSuggestions = [];
    let activeSuggestionIndex = -1;
    let suggestionInputTarget = null;

    let currentSortProperty = 'path';
    let currentSortOrder = 'asc';

    let customRules = [];
    let editingRuleId = null;

    const IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.webp', '.bmp', '.gif'];
    const TAG_EXTENSION = '.txt';
    const THEME_STORAGE_KEY = 'imageTagManagerTheme';
    const CUSTOM_RULES_STORAGE_KEY = 'imageTagManagerCustomRules';

    // --- Helper Functions ---
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    function generateRuleId(prefix = 'rule') {
        return `${prefix}-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;
    }

    function escapeHtml(unsafe) {
        return unsafe
             .replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;")
             .replace(/'/g, "&#039;");
    }

    // --- Theme Management ---
    function applyTheme(theme) {
        if (theme === 'dark') {
            document.body.classList.add('dark-mode');
            themeToggleBtn.textContent = 'Light Mode';
        } else {
            document.body.classList.remove('dark-mode');
            themeToggleBtn.textContent = 'Dark Mode';
        }
    }

    function toggleTheme() {
        const currentTheme = document.body.classList.contains('dark-mode') ? 'light' : 'dark';
        applyTheme(currentTheme);
        try {
            localStorage.setItem(THEME_STORAGE_KEY, currentTheme);
        } catch (e) {
            console.warn("Could not save theme preference:", e);
        }
    }

    function loadInitialTheme() {
        let preferredTheme = 'light';
        try {
            const savedTheme = localStorage.getItem(THEME_STORAGE_KEY);
            if (savedTheme) preferredTheme = savedTheme;
            else if (window.matchMedia?.('(prefers-color-scheme: dark)')?.matches) preferredTheme = 'dark';
        } catch (e) {
            console.warn("Could not access localStorage for theme:", e);
        }
        applyTheme(preferredTheme);
    }

    // --- Core Functions ---
    function updateStatus(message, isLoading = false, isScanning = false) {
        statusMessage.textContent = `Status: ${message}`;
        loadingIndicator.classList.toggle('hidden', !isLoading);
        const progressElement = loadingIndicator.querySelector('progress');
        if (progressElement) {
            progressElement.classList.toggle('hidden', isScanning);
            if (!isScanning) {
                progressElement.value = 0;
                progressElement.removeAttribute('max');
            }
        }
        console.log(message);
    }
    async function processDirectoryRecursively(directoryHandle, fileHandlesMap, currentPath = '') {
        allFileSystemEntries.set(currentPath || '.', { handle: directoryHandle, kind: 'directory', name: directoryHandle.name });
        try {
            for await (const entry of directoryHandle.values()) {
                const entryPath = currentPath ? `${currentPath}/${entry.name}` : entry.name;
                allFileSystemEntries.set(entryPath, { handle: entry, kind: entry.kind, name: entry.name });

                if (entry.kind === 'file') {
                    const lowerCaseName = entry.name.toLowerCase();
                    const ext = lowerCaseName.substring(lowerCaseName.lastIndexOf('.'));
                    const baseNameWithoutExt = lowerCaseName.substring(0, lowerCaseName.lastIndexOf('.'));
                    const mapKey = currentPath ? `${currentPath}/${baseNameWithoutExt}` : baseNameWithoutExt;
                    
                    const originalNameNoExt = entry.name.substring(0, entry.name.lastIndexOf('.'));
                    
                    let pair = fileHandlesMap.get(mapKey) || {
                        imageHandle: null,
                        tagHandle: null,
                        originalImageName: null, // Store original case for image name
                        originalTagName: null,   // Store original case for tag name
                        baseNameKey: mapKey,     // Store the case-insensitive base name key
                        relativePath: currentPath
                    };

                    if (IMAGE_EXTENSIONS.includes(ext)) {
                        pair.imageHandle = entry;
                        pair.originalImageName = entry.name; 
                    } else if (ext === TAG_EXTENSION) {
                        pair.tagHandle = entry;
                        pair.originalTagName = entry.name;
                    }
                    
                    if (IMAGE_EXTENSIONS.includes(ext) || ext === TAG_EXTENSION) {
                         fileHandlesMap.set(mapKey, pair);
                    }

                } else if (entry.kind === 'directory') {
                    await processDirectoryRecursively(entry, fileHandlesMap, entryPath);
                }
            }
        } catch (err) {
            console.warn(`Could not process dir "${currentPath}/${directoryHandle.name}": ${err.message}. Skipping.`);
            updateStatus(`Warning: Could not fully access directory "${currentPath}/${directoryHandle.name}". Some files might be missed.`, false);
        }
    }
    async function loadDataset() {
        try {
            if (!window.showDirectoryPicker) {
                updateStatus("Error: Browser lacks API support.");
                return;
            }
            resetAppState();
            datasetHandle = await window.showDirectoryPicker({
                startIn: "downloads"
            });
            if (!datasetHandle) {
                updateStatus("No directory selected.");
                return;
            }
            updateStatus('Scanning directory structure...', true, true);
            let fileHandlesMap = new Map(); // Key: relativePath/baseName (lowercase, no ext), Value: {imageHandle, tagHandle, originalImageName, originalTagName, relativePath}
            allFileSystemEntries.clear();

            if (await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted' &&
                await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted') {
                updateStatus('Error: Permission denied.');
                datasetHandle = null;
                return;
            }

            await processDirectoryRecursively(datasetHandle, fileHandlesMap, '');
            
            updateStatus('Processing files (tags, metadata, dimensions)...', true, false);
            let processedCount = 0;
            const totalFilesToProcess = fileHandlesMap.size;
            const progressElement = loadingIndicator.querySelector('progress');
            if (progressElement) progressElement.max = totalFilesToProcess;

            for (const [, pair] of fileHandlesMap.entries()) {
                if (pair.imageHandle) { // Prioritize pairs with an image
                    let tags = [];
                    let originalTags = '';
                    let tagHandle = pair.tagHandle;
                    
                    const imageNameForRecord = pair.originalImageName.substring(0, pair.originalImageName.lastIndexOf('.'));

                    if (tagHandle) {
                        try {
                            const file = await tagHandle.getFile();
                            originalTags = await file.text();
                            tags = parseTags(originalTags);
                        } catch (e) {
                            console.warn(`Tag read error for ${imageNameForRecord}: ${e.message}`);
                        }
                    }

                    try {
                        const imageFile = await pair.imageHandle.getFile();
                        const imageUrl = URL.createObjectURL(imageFile);
                        const dimensions = await new Promise((resolve) => {
                            const img = new Image();
                            img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
                            img.onerror = () => {
                                console.warn(`Could not load image ${imageNameForRecord} to get dimensions.`);
                                resolve({ width: 0, height: 0 });
                            };
                            img.src = imageUrl;
                        });

                        allImageData.push({
                            imageHandle: pair.imageHandle,
                            tagHandle: tagHandle,
                            imageName: imageNameForRecord, // Base name without extension, original case from image
                            originalFullImageName: pair.originalImageName, // Full original image filename
                            originalFullTagName: pair.originalTagName, // Full original tag filename (if exists)
                            relativePath: pair.relativePath,
                            imageUrl: imageUrl,
                            tags: tags,
                            originalTags: originalTags,
                            modified: false,
                            id: `item-${nextItemId++}`,
                            fileSize: imageFile.size,
                            lastModified: imageFile.lastModified,
                            imageWidth: dimensions.width,
                            imageHeight: dimensions.height,
                        });
                        updateTagStats(tags, [], true);
                    } catch (e) {
                        console.warn(`Image processing error for ${imageNameForRecord}: ${e.message}`);
                    }
                }
                processedCount++;
                if (progressElement) progressElement.value = processedCount;
            }

            initialTotalCount = allImageData.length;
            displayedImageDataIndices = allImageData.map((_, i) => i);
            currentSortProperty = sortPropertySelect.value;
            currentSortOrder = sortOrderSelect.value;
            sortAndRender();
            updateStatisticsDisplay();
            saveAllBtn.disabled = true;
        } catch (error) {
            if (error.name === 'AbortError') updateStatus('Dataset loading cancelled.');
            else updateStatus(`Error loading dataset: ${error.message}`);
            console.error(error);
            resetAppState();
            datasetHandle = null;
        } finally {
            loadingIndicator.classList.add('hidden');
        }
    }

    function resetAppState() {
        allImageData.forEach(item => {
            if (item.imageUrl) URL.revokeObjectURL(item.imageUrl);
        });
        allImageData = [];
        allFileSystemEntries.clear();
        displayedImageDataIndices = [];
        tagFrequencies.clear();
        uniqueTags.clear();
        imageGrid.innerHTML = '';
        tagFrequencySearchInput.value = '';
        updateStatisticsDisplay();
        updateDisplayedCount();
        saveAllBtn.disabled = true;
        duplicateFilesOutput.innerHTML = '';
        orphanedMissingOutput.innerHTML = '';
        caseInconsistentOutput.innerHTML = '';
        currentEditIndex = -1;
        nextItemId = 0;
        initialTotalCount = 0;
        currentSortProperty = 'path';
        sortPropertySelect.value = currentSortProperty;
        currentSortOrder = 'asc';
        sortOrderSelect.value = currentSortOrder;
        clearKeyboardFocus();
        clearSelection();
        hideSuggestions();
    }

    function parseTags(tagString) {
        if (!tagString || typeof tagString !== 'string') return [];
        return tagString.split(',').map(tag => tag.trim()).filter(tag => tag.length > 0);
    }

    function joinTags(tagArray) {
        return tagArray.map(t => t.trim()).filter(t => t).join(', ');
    }

    function clearKeyboardFocus() {
        if (keyboardFocusIndex !== -1 && displayedImageDataIndices.length > 0 && allImageData[displayedImageDataIndices[keyboardFocusIndex]]) {
            const oldItemId = allImageData[displayedImageDataIndices[keyboardFocusIndex]]?.id;
            if (oldItemId) {
                document.getElementById(oldItemId)?.classList.remove('keyboard-focus');
            }
        }
        keyboardFocusIndex = -1;
    }

    function setKeyboardFocus(newIndex) {
        if (newIndex < 0 || newIndex >= displayedImageDataIndices.length) return;
        clearKeyboardFocus();
        keyboardFocusIndex = newIndex;
        const currentItemId = allImageData[displayedImageDataIndices[keyboardFocusIndex]]?.id;
        if (currentItemId) {
            const element = document.getElementById(currentItemId);
            if (element) {
                element.classList.add('keyboard-focus');
                element.scrollIntoView({
                    block: 'nearest',
                    inline: 'nearest'
                });
            }
        }
    }

    function renderImageGrid() {
        clearKeyboardFocus();
        imageGrid.innerHTML = '';
        const fragment = document.createDocumentFragment();
        displayedImageDataIndices.forEach((globalIndex, displayIndex) => {
            const item = allImageData[globalIndex];
            if (!item) return;
            const div = document.createElement('div');
            div.classList.add('grid-item');
            div.dataset.globalIndex = globalIndex;
            div.dataset.displayIndex = displayIndex;
            div.id = item.id;
            if (item.modified) div.classList.add('modified');
            if (selectedItemIds.has(item.id)) div.classList.add('selected');
            div.tabIndex = -1;
            div.style.outline = 'none';
            div.addEventListener('click', (e) => handleGridItemInteraction(e, item.id, globalIndex, displayIndex));
            const imgDeleteBtn = document.createElement('span');
            imgDeleteBtn.classList.add('image-delete-btn');
            imgDeleteBtn.textContent = '×';
            const deleteTitle = item.relativePath ? `Remove image "${item.relativePath}/${item.originalFullImageName}" (from view)` : `Remove image "${item.originalFullImageName}" (from view)`;
            imgDeleteBtn.title = deleteTitle;
            imgDeleteBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                requestImageDeletion(item.id);
            });
            div.appendChild(imgDeleteBtn);
            const img = document.createElement('img');
            img.src = item.imageUrl;
            const imgAlt = item.relativePath ? `${item.relativePath}/${item.originalFullImageName}` : item.originalFullImageName;
            img.alt = imgAlt;
            img.title = `${imgAlt}\nSize: ${item.imageWidth}x${item.imageHeight}\nModified: ${new Date(item.lastModified).toLocaleDateString()}`;
            img.loading = 'lazy';
            div.appendChild(img);
            const tagsDiv = document.createElement('div');
            tagsDiv.classList.add('tags');
            renderTagsForItem(tagsDiv, item);
            div.appendChild(tagsDiv);
            fragment.appendChild(div);
        });
        imageGrid.appendChild(fragment);
        updateDisplayedCount();
        updateBatchEditButtonState();
    }

    function handleGridItemInteraction(event, itemId, globalItemIndex, displayItemIndex) {
        const isImageClick = event.target.tagName === 'IMG';
        if (event.shiftKey && lastInteractedItemId && lastInteractedItemId !== itemId) {
            const lastInteractedGlobalIndex = findIndexById(lastInteractedItemId);
            if (lastInteractedGlobalIndex === -1) {
                lastInteractedItemId = null;
                const wasSelected = selectedItemIds.has(itemId);
                const wasOnlySelection = wasSelected && selectedItemIds.size === 1;
                if (wasOnlySelection) {
                    selectedItemIds.delete(itemId);
                    document.getElementById(itemId)?.classList.remove('selected');
                    lastInteractedItemId = null;
                } else {
                    clearSelection();
                    addSelection(itemId);
                    lastInteractedItemId = itemId;
                }
            } else {
                const lastInteractedDisplayIndex = displayedImageDataIndices.indexOf(lastInteractedGlobalIndex);
                const currentDisplayIndex = displayItemIndex;
                if (lastInteractedDisplayIndex === -1) {
                    const wasSelected = selectedItemIds.has(itemId);
                    const wasOnlySelection = wasSelected && selectedItemIds.size === 1;
                    if (wasOnlySelection) {
                        selectedItemIds.delete(itemId);
                        document.getElementById(itemId)?.classList.remove('selected');
                        lastInteractedItemId = null;
                    } else {
                        clearSelection();
                        addSelection(itemId);
                        lastInteractedItemId = itemId;
                    }
                } else {
                    const start = Math.min(lastInteractedDisplayIndex, currentDisplayIndex);
                    const end = Math.max(lastInteractedDisplayIndex, currentDisplayIndex);
                    if (!(event.ctrlKey || event.metaKey)) {
                        selectedItemIds.forEach(id => document.getElementById(id)?.classList.remove('selected'));
                        selectedItemIds.clear();
                    }
                    for (let i = start; i <= end; i++) {
                        const gIdx = displayedImageDataIndices[i];
                        if (allImageData[gIdx]) {
                            addSelection(allImageData[gIdx].id);
                        }
                    }
                }
            }
        } else if (event.ctrlKey || event.metaKey) {
            toggleSelection(itemId);
            lastInteractedItemId = itemId;
        } else {
            const wasSelected = selectedItemIds.has(itemId);
            const wasOnlySelection = wasSelected && selectedItemIds.size === 1;
            if (wasOnlySelection) {
                selectedItemIds.delete(itemId);
                document.getElementById(itemId)?.classList.remove('selected');
                lastInteractedItemId = null;
            } else {
                clearSelection();
                addSelection(itemId);
                lastInteractedItemId = itemId;
            }
            if (isImageClick && selectedItemIds.has(itemId) && selectedItemIds.size === 1) {
                openEditor(itemId);
            }
        }
        updateBatchEditButtonState();
    }

    function toggleSelection(itemId) {
        const element = document.getElementById(itemId);
        if (selectedItemIds.has(itemId)) {
            selectedItemIds.delete(itemId);
            element?.classList.remove('selected');
        } else {
            selectedItemIds.add(itemId);
            element?.classList.add('selected');
        }
    }

    function addSelection(itemId) {
        const element = document.getElementById(itemId);
        if (!selectedItemIds.has(itemId)) {
            selectedItemIds.add(itemId);
            element?.classList.add('selected');
        }
    }

    function clearSelection() {
        lastInteractedItemId = null;
        selectedItemIds.forEach(id => {
            document.getElementById(id)?.classList.remove('selected');
        });
        selectedItemIds.clear();
        updateBatchEditButtonState();
    }

    function updateBatchEditButtonState() {
        const count = selectedItemIds.size;
        batchEditSelectionBtn.textContent = `Edit Selected (${count})`;
        batchEditSelectionBtn.disabled = count === 0;
    }

    function renderTagsForItem(tagsContainer, item) {
        tagsContainer.innerHTML = '';
        if (item.tags.length === 0) {
            tagsContainer.textContent = '(no tags)';
        } else {
            item.tags.forEach(tag => {
                const tagSpan = document.createElement('span');
                tagSpan.classList.add('tag');
                tagSpan.textContent = tag;
                const deleteBtn = document.createElement('span');
                deleteBtn.classList.add('tag-delete-btn');
                deleteBtn.textContent = '×';
                deleteBtn.title = `Remove tag "${tag}"`;
                deleteBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    removeTagFromImage(item.id, tag);
                });
                tagSpan.appendChild(deleteBtn);
                tagsContainer.appendChild(tagSpan);
            });
        }
    }

    function updateDisplayedCount() {
        displayedCountSpan.textContent = displayedImageDataIndices.length;
        totalCountSpan.textContent = allImageData.length;
    }

    function recalculateAllTagStats() {
        tagFrequencies.clear();
        uniqueTags.clear();
        allImageData.forEach(item => {
            item.tags.forEach(tag => {
                const t = tag.trim();
                if (t) {
                    tagFrequencies.set(t, (tagFrequencies.get(t) || 0) + 1);
                    uniqueTags.add(t);
                }
            });
        });
        updateStatisticsDisplay();
    }

    function updateTagStatsIncremental(addedTags, removedTags) {
        const add = addedTags.map(t => t.trim()).filter(t => t);
        const rem = removedTags.map(t => t.trim()).filter(t => t);
        rem.forEach(tag => {
            if (tagFrequencies.has(tag)) {
                const count = tagFrequencies.get(tag) - 1;
                if (count <= 0) {
                    tagFrequencies.delete(tag);
                    uniqueTags.delete(tag);
                } else {
                    tagFrequencies.set(tag, count);
                }
            }
        });
        add.forEach(tag => {
            tagFrequencies.set(tag, (tagFrequencies.get(tag) || 0) + 1);
            uniqueTags.add(tag);
        });
    }

    function updateTagStats(addedTags, removedTags, isInitialLoad = false) {
        if (isInitialLoad) {
            addedTags.map(t => t.trim()).filter(t => t).forEach(tag => {
                tagFrequencies.set(tag, (tagFrequencies.get(tag) || 0) + 1);
                uniqueTags.add(tag);
            });
        } else {
            updateTagStatsIncremental(addedTags, removedTags);
        }
    }

    function updateStatisticsDisplay() {
        statsTotalImages.textContent = allImageData.length;
        statsUniqueTags.textContent = uniqueTags.size;
        statsModifiedImages.textContent = allImageData.filter(item => item.modified).length;
        tagFrequencyList.innerHTML = '';
        const searchTerm = tagFrequencySearchInput.value.trim().toLowerCase();
        const filteredFrequencies = [...tagFrequencies.entries()].filter(([t]) => !searchTerm || t.toLowerCase().includes(searchTerm));
        const sortedTags = filteredFrequencies.sort((a, b) => (b[1] - a[1]) || a[0].localeCompare(b[0]));
        if (sortedTags.length === 0) {
            tagFrequencyList.innerHTML = `<div><span>${searchTerm ? 'No tags match search.' : 'No tags loaded.'}</span></div>`;
        } else {
            sortedTags.forEach(([tag, count]) => {
                const div = document.createElement('div');
                const spanT = document.createElement('span');
                spanT.textContent = tag;
                spanT.title = `Click to search for: ${tag}`;
                spanT.addEventListener('click', () => {
                    const currentSearch = searchTagsInput.value.trim();
                    if (currentSearch) {
                        searchTagsInput.value = `${currentSearch}, ${tag}`;
                    } else {
                        searchTagsInput.value = tag;
                    }
                    filterBtn.click();
                    searchTagsInput.focus();
                });
                const spanC = document.createElement('span');
                spanC.textContent = count;
                div.appendChild(spanT);
                div.appendChild(spanC);
                tagFrequencyList.appendChild(div);
            });
        }
    }

    function sortAndRender() {
        clearKeyboardFocus();
        displayedImageDataIndices.sort((indexA, indexB) => {
            const itemA = allImageData[indexA];
            const itemB = allImageData[indexB];
            let comparison = 0;
            if (!itemA || !itemB) return 0;
            switch (currentSortProperty) {
                case 'filename':
                    comparison = (itemA.originalFullImageName || '').localeCompare(itemB.originalFullImageName || '');
                    break;
                case 'tagcount':
                    comparison = (itemA.tags?.length || 0) - (itemB.tags?.length || 0);
                    break;
                case 'lastmodified':
                    comparison = (itemA.lastModified || 0) - (itemB.lastModified || 0);
                    break;
                case 'filesize':
                    comparison = (itemA.fileSize || 0) - (itemB.fileSize || 0);
                    break;
                case 'dimensions':
                    const areaA = (itemA.imageWidth || 0) * (itemA.imageHeight || 0);
                    const areaB = (itemB.imageWidth || 0) * (itemB.imageHeight || 0);
                    comparison = areaA - areaB;
                    break;
                case 'width':
                    comparison = (itemA.imageWidth || 0) - (itemB.imageWidth || 0);
                    break;
                case 'height':
                    comparison = (itemA.imageHeight || 0) - (itemB.imageHeight || 0);
                    break;
                case 'path':
                default:
                    const pathA = `${itemA.relativePath || ''}/${itemA.originalFullImageName || ''}`;
                    const pathB = `${itemB.relativePath || ''}/${itemB.originalFullImageName || ''}`;
                    comparison = pathA.localeCompare(pathB);
                    break;
            }
            return currentSortOrder === 'asc' ? comparison : -comparison;
        });
        renderImageGrid();
        updateStatus(`Displaying ${displayedImageDataIndices.length} images. (Sorted by ${currentSortProperty} ${currentSortOrder})`, false);
    }

    function filterAndSearch() {
        clearSelection();
        const searchTermsRaw = parseTags(searchTagsInput.value);
        const excludeTerms = parseTags(excludeTagsInput.value);
        const searchMode = searchModeSelect.value;
        const filterMode = filterUntaggedSelect.value;
        updateStatus('Filtering...', true, true);
        requestAnimationFrame(() => {
            displayedImageDataIndices = allImageData.map((_, i) => i).filter(index => {
                const item = allImageData[index];
                if (!item) return false;
                let include = true;
                switch (filterMode) {
                    case 'tagged':
                        if (item.tags.length === 0) include = false;
                        break;
                    case 'untagged':
                        if (item.tags.length > 0) include = false;
                        break;
                    case 'modified':
                        if (!item.modified) include = false;
                        break;
                }
                if (!include) return false;
                if (searchTermsRaw.length > 0) {
                    const itemTagsLower = item.tags.map(t => t.toLowerCase());
                    if (searchMode === 'AND') {
                        include = searchTermsRaw.every(term => itemTagsLower.some(tag => tag.includes(term.toLowerCase())));
                    } else {
                        include = searchTermsRaw.some(term => itemTagsLower.some(tag => tag.includes(term.toLowerCase())));
                    }
                }
                if (!include) return false;
                if (excludeTerms.length > 0) {
                    include = !excludeTerms.some(term => item.tags.includes(term));
                }
                return include;
            });
            sortAndRender();
        });
    }

    function resetFilters() {
        clearSelection();
        searchTagsInput.value = '';
        excludeTagsInput.value = '';
        searchModeSelect.value = 'AND';
        filterUntaggedSelect.value = 'all';
        duplicateFilesOutput.innerHTML = '';
        orphanedMissingOutput.innerHTML = '';
        caseInconsistentOutput.innerHTML = '';
        currentSortProperty = 'path';
        sortPropertySelect.value = currentSortProperty;
        currentSortOrder = 'asc';
        sortOrderSelect.value = currentSortOrder;
        displayedImageDataIndices = allImageData.map((_, i) => i);
        sortAndRender();
    }

    function findIndexById(itemId) {
        return allImageData.findIndex(item => item.id === itemId);
    }

    function openEditor(itemId) {
        const index = findIndexById(itemId);
        if (index === -1) {
            updateStatus("Error: Item not found for single editor.");
            return;
        }
        currentEditIndex = index;
        const item = allImageData[index];
        currentEditOriginalTagsString = joinTags(item.tags);
        editorImagePreview.src = item.imageUrl;
        if (editorImagePreview.classList.contains('zoomed')) {
            toggleEditorImageZoom();
        }
        
        const tagFileName = item.originalFullTagName || `${item.imageName}${TAG_EXTENSION}`;
        const editorDisplayName = item.relativePath ? `${item.relativePath}/${tagFileName}` : tagFileName;
        editorFilename.textContent = editorDisplayName;

        let metadataHTML = `Dimensions: ${item.imageWidth} x ${item.imageHeight}px`;
        metadataHTML += `  |  Size: ${formatFileSize(item.fileSize)}`;
        metadataHTML += `  |  Modified: ${new Date(item.lastModified).toLocaleDateString()}`;
        editorMetadataDisplay.innerHTML = metadataHTML;
        editorTagsTextarea.value = joinTags(item.tags);
        editorStatus.textContent = '';
        editorModal.classList.remove('hidden');
        updateEditorNavButtonStates();
        hideSuggestions();
    }

    function closeEditor() {
        if (editorImagePreview.classList.contains('zoomed')) {
            toggleEditorImageZoom();
        }
        editorModal.classList.add('hidden');
        currentEditIndex = -1;
        currentEditOriginalTagsString = '';
        editorMetadataDisplay.innerHTML = '';
        if (keyboardFocusIndex !== -1 && displayedImageDataIndices.length > keyboardFocusIndex) {
            const focusedItemId = allImageData[displayedImageDataIndices[keyboardFocusIndex]]?.id;
            document.getElementById(focusedItemId)?.focus();
        }
        hideSuggestions();
    }

    function saveEditorChanges() {
        if (currentEditIndex === -1 || !allImageData[currentEditIndex]) {
            editorStatus.textContent = "Error: Data missing.";
            editorStatus.style.color = 'var(--error-color)';
            return false;
        }
        const item = allImageData[currentEditIndex];
        const newTagsFromTextarea = parseTags(editorTagsTextarea.value);
        const newTagsStringFromTextarea = joinTags(newTagsFromTextarea);
        if (newTagsStringFromTextarea !== currentEditOriginalTagsString) {
            item.tags = newTagsFromTextarea;
            item.modified = true;
            saveAllBtn.disabled = false;
            const oldTagsForStatUpdate = parseTags(currentEditOriginalTagsString);
            const addedForStatUpdate = newTagsFromTextarea.filter(t => !oldTagsForStatUpdate.includes(t));
            const removedForStatUpdate = oldTagsForStatUpdate.filter(t => !newTagsFromTextarea.includes(t));
            updateTagStatsIncremental(addedForStatUpdate, removedForStatUpdate);
            recalculateAllTagStats();
            const gridItem = document.getElementById(item.id);
            if (gridItem) {
                gridItem.classList.add('modified');
                const tagsDiv = gridItem.querySelector('.tags');
                if (tagsDiv) renderTagsForItem(tagsDiv, item);
            }
            editorStatus.textContent = 'Tags updated. Save changes.';
            editorStatus.style.color = 'var(--success-color)';
            currentEditOriginalTagsString = newTagsStringFromTextarea;
            return true;
        } else {
            editorStatus.textContent = 'No changes detected to save.';
            editorStatus.style.color = 'var(--status-color)';
            return false;
        }
    }

    function removeDuplicateTagsInEditor() {
        if (currentEditIndex === -1 || !allImageData[currentEditIndex]) return;
        const currentTags = parseTags(editorTagsTextarea.value);
        const unique = [...new Set(currentTags)];
        editorTagsTextarea.value = joinTags(unique);
        const removedCount = currentTags.length - unique.length;
        if (removedCount > 0) {
            editorStatus.textContent = `Removed ${removedCount} duplicates.`;
            editorStatus.style.color = 'var(--success-color)';
        } else {
            editorStatus.textContent = 'No duplicates found.';
            editorStatus.style.color = 'var(--status-color)';
        }
    }

    function removeTagFromImage(itemId, tagToRemove) {
        const index = findIndexById(itemId);
        if (index === -1) return;
        const item = allImageData[index];
        const oldLen = item.tags.length;
        const newTags = item.tags.filter(t => t !== tagToRemove);
        if (newTags.length < oldLen) {
            item.tags = newTags;
            item.modified = true;
            saveAllBtn.disabled = false;
            updateTagStatsIncremental([], [tagToRemove]);
            recalculateAllTagStats();
            const gridItem = document.getElementById(item.id);
            if (gridItem) {
                gridItem.classList.add('modified');
                const tagsDiv = gridItem.querySelector('.tags');
                if (tagsDiv) renderTagsForItem(tagsDiv, item);
            }
        }
    }

    function requestImageDeletion(itemId) {
        const index = findIndexById(itemId);
        if (index === -1) {
            updateStatus("Error: Item not found.");
            return;
        }
        const item = allImageData[index];
        const deleteMsg = item.relativePath ? `Permanently DELETE "${item.relativePath}/${item.originalFullImageName}" and its tag file?` : `Permanently DELETE "${item.originalFullImageName}" and its tag file?`;
        if (confirm(`${deleteMsg}\n\n*** WARNING: PERMANENT! ***`)) {
            deleteImageFilesAndView(itemId);
        } else {
            updateStatus("Deletion cancelled.");
        }
    }
    async function deleteImageFilesAndView(itemId) {
        const index = findIndexById(itemId);
        if (index === -1) return;
        const item = allImageData[index];
        const fullImagePath = item.relativePath ? `${item.relativePath}/${item.originalFullImageName}` : item.originalFullImageName;
        let errorOccurred = false;
        updateStatus(`Attempting to delete ${fullImagePath}...`, true);
        try {
            if (!datasetHandle || await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted' && await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted') {
                updateStatus('Error: Permission denied.');
                return;
            }
            let parentDirHandle = datasetHandle;
            if (item.relativePath) {
                try {
                    const parts = item.relativePath.split('/');
                    for (const p of parts) {
                        if (p) parentDirHandle = await parentDirHandle.getDirectoryHandle(p, { create: false });
                    }
                } catch (e) {
                    console.error(`Dir nav error:`, e);
                    updateStatus(`Error finding dir "${item.relativePath}"`);
                    errorOccurred = true;
                    item.imageHandle = null;
                    item.tagHandle = null;
                }
            }
            if (item.imageHandle && !errorOccurred) {
                const name = item.imageHandle.name; // Use the handle's name for exact match
                try {
                    await parentDirHandle.removeEntry(name);
                    console.log(`Deleted image: ${fullImagePath}`);
                    item.imageHandle = null;
                } catch (e) {
                    console.error(`Img delete error:`, e);
                    updateStatus(`Error deleting image "${fullImagePath}"`);
                    errorOccurred = true;
                }
            }
            if (item.tagHandle && !errorOccurred) {
                const name = item.tagHandle.name; // Use the handle's name
                const fullTagPath = item.relativePath ? `${item.relativePath}/${name}` : name;
                try {
                    await parentDirHandle.removeEntry(name);
                    console.log(`Deleted tag file for: ${fullTagPath}`);
                    item.tagHandle = null;
                } catch (e) {
                    console.error(`Tag delete error:`, e);
                    if (!errorOccurred) updateStatus(`Error deleting tag file "${fullTagPath}"`);
                    errorOccurred = true;
                }
            }
            document.getElementById(itemId)?.remove();
            if (item.imageUrl) URL.revokeObjectURL(item.imageUrl);
            if (selectedItemIds.has(itemId)) {
                selectedItemIds.delete(itemId);
                updateBatchEditButtonState();
            }
            if (lastInteractedItemId === itemId) lastInteractedItemId = null;
            const originalGlobalIndexForDisplayed = displayedImageDataIndices.indexOf(index);
            allImageData.splice(index, 1);
            displayedImageDataIndices = displayedImageDataIndices.reduce((acc, dispIdx) => {
                if (dispIdx < index) acc.push(dispIdx);
                else if (dispIdx > index) acc.push(dispIdx - 1);
                return acc;
            }, []);
            recalculateAllTagStats();
            updateDisplayedCount();
            if (!errorOccurred) updateStatus(`Deleted "${fullImagePath}" files. ${allImageData.length} images left.`);
            else updateStatus(statusMessage.textContent + ` Removed from view despite errors.`);
            if (keyboardFocusIndex === originalGlobalIndexForDisplayed && originalGlobalIndexForDisplayed !== -1) {
                clearKeyboardFocus();
            } else if (keyboardFocusIndex > originalGlobalIndexForDisplayed && originalGlobalIndexForDisplayed !== -1) {
                setKeyboardFocus(keyboardFocusIndex - 1);
            }
            if (displayedImageDataIndices.length === 0 && allImageData.length > 0) {
                resetFilters();
            } else if (displayedImageDataIndices.length > 0) {
                renderImageGrid(); // Rerender might not be strictly necessary but good for consistency
            }
        } catch (permError) {
            console.error("Permission error:", permError);
            updateStatus(`Perm error: ${permError.message}`);
        } finally {
            loadingIndicator.classList.add('hidden');
        }
    }

    // --- Bulk Operations (Standard) ---
    function performStandardBulkOperation(operation) {
        if (!datasetHandle || allImageData.length === 0) {
            updateStatus("Load dataset first.");
            return;
        }
        let modCount = 0;
        let affectedIds = new Set();
        let oldTag, newTag, toRemove, toAdd;
        let confirmNeeded = false;
        let desc = "";
        switch (operation) {
            case 'rename':
                oldTag = renameOldTagInput.value.trim();
                newTag = renameNewTagInput.value.trim();
                if (!oldTag || !newTag || oldTag === newTag) {
                    updateStatus("Invalid rename input.");
                    return;
                }
                desc = `Rename "${oldTag}" to "${newTag}"`;
                break;
            case 'remove':
                toRemove = removeTagNameInput.value.trim();
                if (!toRemove) {
                    updateStatus("Need tag to remove.");
                    return;
                }
                desc = `Remove "${toRemove}"`;
                confirmNeeded = true;
                break;
            case 'add':
                toAdd = parseTags(addTriggerWordsInput.value);
                if (toAdd.length === 0) {
                    updateStatus("Need tags to add.");
                    return;
                }
                desc = `Add tags "${joinTags(toAdd)}"`;
                break;
            case 'removeDuplicates':
                desc = `Remove duplicate tags from each file`;
                break;
            default:
                updateStatus("Unknown standard bulk op.");
                return;
        }
        if (confirmNeeded && !confirm(`Bulk action on ALL images?\n\n${desc}`)) {
            updateStatus("Cancelled.");
            return;
        }
        updateStatus(`${desc}...`, true);
        let statsChanged = false;
        allImageData.forEach((item) => {
            let origTagsString = joinTags(item.tags); 
            let currentTagsArray = [...item.tags];
            let itemChanged = false;
            switch (operation) {
                case 'rename':
                    if (currentTagsArray.includes(oldTag)) {
                        currentTagsArray = [...new Set(currentTagsArray.map(t => t === oldTag ? newTag : t))];
                        itemChanged = true;
                    }
                    break;
                case 'remove':
                    const initialLength = currentTagsArray.length;
                    currentTagsArray = currentTagsArray.filter(t => t !== toRemove);
                    if (currentTagsArray.length < initialLength) itemChanged = true;
                    break;
                case 'add':
                    const uniqueTagsToAdd = toAdd.filter(t => !currentTagsArray.includes(t));
                    if (uniqueTagsToAdd.length > 0) {
                        currentTagsArray = [...uniqueTagsToAdd, ...currentTagsArray];
                        itemChanged = true;
                    }
                    break;
                case 'removeDuplicates':
                    const uniqueTagsArr = [...new Set(currentTagsArray)];
                    if (uniqueTagsArr.length < currentTagsArray.length) {
                        currentTagsArray = uniqueTagsArr;
                        itemChanged = true;
                    }
                    break;
            }
            if (itemChanged && joinTags(currentTagsArray) === origTagsString && operation !== 'removeDuplicates') {
                itemChanged = false; 
            }

            if (itemChanged) {
                item.tags = currentTagsArray;
                item.modified = true;
                modCount++;
                affectedIds.add(item.id);
                statsChanged = true;
            }
        });
        if (modCount > 0) {
            saveAllBtn.disabled = false;
            affectedIds.forEach(id => {
                const gridItem = document.getElementById(id);
                const itemData = allImageData.find(i => i.id === id);
                if (gridItem && itemData) {
                    gridItem.classList.add('modified');
                    const tagsDiv = gridItem.querySelector('.tags');
                    if (tagsDiv) renderTagsForItem(tagsDiv, itemData);
                }
            });
            if (statsChanged) recalculateAllTagStats();
            updateStatus(`${modCount} items affected by "${desc}". Save changes.`, false);
        } else {
            updateStatus(`Bulk op "${desc}" complete. No changes.`, false);
        }
        if (operation === 'rename') {
            renameOldTagInput.value = '';
            renameNewTagInput.value = '';
        }
        if (operation === 'remove') removeTagNameInput.value = '';
        if (operation === 'add') addTriggerWordsInput.value = '';
    }

    // --- UTILITY FUNCTIONS ---
    function findDuplicateTagFiles() {
        if (allImageData.length < 2) {
            duplicateFilesOutput.innerHTML = '<p>Need at least 2 files to compare.</p>';
            return;
        }
        updateStatus('Finding files with identical tags...', true);
        const tagStringToFilesMap = new Map();
        allImageData.forEach(item => {
            const tagString = joinTags(item.tags.slice().sort());
            const displayPath = item.relativePath ? `${item.relativePath}/${item.originalFullImageName}` : item.originalFullImageName;
            if (!tagStringToFilesMap.has(tagString)) {
                tagStringToFilesMap.set(tagString, []);
            }
            tagStringToFilesMap.get(tagString).push(displayPath);
        });
        const duplicateSets = [];
        for (const [tagStr, files] of tagStringToFilesMap.entries()) {
            if (files.length > 1) {
                duplicateSets.push({
                    tags: tagStr,
                    files: files.sort()
                });
            }
        }
        duplicateSets.sort((a, b) => a.tags.localeCompare(b.tags));
        duplicateFilesOutput.innerHTML = '';
        if (duplicateSets.length > 0) {
            const summaryP = document.createElement('p');
            summaryP.textContent = `Found ${duplicateSets.length} set(s) of files with identical tags:`;
            duplicateFilesOutput.appendChild(summaryP);
            duplicateSets.forEach(dupSet => {
                const p = document.createElement('p');
                p.textContent = `Tags ${dupSet.tags ? `"${escapeHtml(dupSet.tags)}"` : '(no tags)'}:`;
                const ul = document.createElement('ul');
                dupSet.files.forEach(fileName => {
                    const li = document.createElement('li');
                    li.textContent = escapeHtml(fileName);
                    ul.appendChild(li);
                });
                duplicateFilesOutput.appendChild(p);
                duplicateFilesOutput.appendChild(ul);
            });
            updateStatus(`Found ${duplicateSets.length} duplicate set(s).`, false);
        } else {
            duplicateFilesOutput.innerHTML = '<p>No files found with identical tags.</p>';
            updateStatus('No duplicate tag sets found.', false);
        }
    }

    async function findOrphanedAndMissingTagFiles() {
        if (!datasetHandle) {
            orphanedMissingOutput.innerHTML = '<p>Load dataset first.</p>';
            return;
        }
        updateStatus('Checking for orphaned/missing tag files...', true);
        const orphanedTagFiles = [];
        const imagesMissingTagFiles = [];
        const allImageBaseNames = new Set(); // Stores lowercased base names of images from allImageData
        const allTagFileBaseNames = new Set(); // Stores lowercased base names of tag files from allImageData

        // Populate sets from allImageData (our primary source of truth for "managed" pairs)
        allImageData.forEach(item => {
            const imageBase = item.imageName.toLowerCase(); // item.imageName is base name
            const itemFullPath = item.relativePath ? `${item.relativePath}/${imageBase}` : imageBase;
            allImageBaseNames.add(itemFullPath);

            if (item.tagHandle) {
                 // item.originalFullTagName will be like "name.txt", need base name
                const tagBase = item.originalFullTagName.substring(0, item.originalFullTagName.lastIndexOf('.')).toLowerCase();
                const tagFullPath = item.relativePath ? `${item.relativePath}/${tagBase}` : tagBase;
                allTagFileBaseNames.add(tagFullPath);
            }
        });
        
        // Iterate through all known file system entries
        for (const [fullPath, entryInfo] of allFileSystemEntries.entries()) {
            if (entryInfo.kind !== 'file') continue;

            const lowerCaseName = entryInfo.name.toLowerCase();
            const ext = lowerCaseName.substring(lowerCaseName.lastIndexOf('.'));
            const baseName = lowerCaseName.substring(0, lowerCaseName.lastIndexOf('.'));
            const pathParts = fullPath.split('/');
            pathParts.pop(); // remove filename
            const relativePath = pathParts.join('/');
            const mapKey = relativePath ? `${relativePath}/${baseName}` : baseName;

            if (IMAGE_EXTENSIONS.includes(ext)) {
                // This is an image file according to file system.
                // Does it have a corresponding tag file in allTagFileBaseNames?
                if (!allTagFileBaseNames.has(mapKey)) {
                     // Also check if it's even in our allImageData. If not, it's an image we didn't load/pair.
                    let foundInAllImageData = false;
                    for(const imgData of allImageData){
                        const imgDataBase = imgData.imageName.toLowerCase();
                        const imgDataPath = imgData.relativePath ? `${imgData.relativePath}/${imgDataBase}` : imgDataBase;
                        if(imgDataPath === mapKey){
                            foundInAllImageData = true;
                            break;
                        }
                    }
                    if(foundInAllImageData){ // Only report if it's an image we manage
                        imagesMissingTagFiles.push(fullPath);
                    }
                }
            } else if (ext === TAG_EXTENSION) {
                // This is a tag file according to file system.
                // Does it have a corresponding image in allImageBaseNames?
                if (!allImageBaseNames.has(mapKey)) {
                    orphanedTagFiles.push(fullPath);
                }
            }
        }


        orphanedMissingOutput.innerHTML = '';
        let foundAny = false;

        if (imagesMissingTagFiles.length > 0) {
            foundAny = true;
            const p = document.createElement('p');
            p.textContent = `Images missing tag files (${imagesMissingTagFiles.length}):`;
            orphanedMissingOutput.appendChild(p);
            const ul = document.createElement('ul');
            imagesMissingTagFiles.sort().forEach(filePath => {
                const li = document.createElement('li');
                li.textContent = escapeHtml(filePath);
                ul.appendChild(li);
            });
            orphanedMissingOutput.appendChild(ul);
        }

        if (orphanedTagFiles.length > 0) {
            foundAny = true;
            const p = document.createElement('p');
            p.textContent = `Orphaned tag files (no matching image) (${orphanedTagFiles.length}):`;
            orphanedMissingOutput.appendChild(p);
            const ul = document.createElement('ul');
            orphanedTagFiles.sort().forEach(filePath => {
                const li = document.createElement('li');
                li.textContent = escapeHtml(filePath);
                ul.appendChild(li);
            });
            orphanedMissingOutput.appendChild(ul);
        }

        if (!foundAny) {
            orphanedMissingOutput.innerHTML = '<p>No orphaned or missing tag files found.</p>';
        }
        updateStatus('Orphaned/missing tag file check complete.', false);
    }

    function findCaseInconsistentTags() {
        if (uniqueTags.size < 2) {
            caseInconsistentOutput.innerHTML = '<p>Not enough unique tags to compare.</p>';
            return;
        }
        updateStatus('Finding case-inconsistent tags...', true);
        const tagMap = new Map(); // lowercase_tag -> [OriginalCaseTag1, OriginalCaseTag2]
        uniqueTags.forEach(tag => {
            const lower = tag.toLowerCase();
            if (!tagMap.has(lower)) {
                tagMap.set(lower, []);
            }
            tagMap.get(lower).push(tag);
        });

        const inconsistentGroups = [];
        for (const [lower, originals] of tagMap.entries()) {
            if (originals.length > 1) {
                 // Further check: ensure there's actually a case difference
                const uniqueOriginals = new Set(originals);
                if (uniqueOriginals.size > 1) { // If "tag", "tag" was present, Set makes it 1
                    inconsistentGroups.push({ lower, originals: [...uniqueOriginals].sort() });
                }
            }
        }
        inconsistentGroups.sort((a,b) => a.lower.localeCompare(b.lower));

        caseInconsistentOutput.innerHTML = '';
        if (inconsistentGroups.length > 0) {
            const summaryP = document.createElement('p');
            summaryP.textContent = `Found ${inconsistentGroups.length} group(s) of case-inconsistent tags:`;
            caseInconsistentOutput.appendChild(summaryP);
            inconsistentGroups.forEach(group => {
                const div = document.createElement('div');
                div.classList.add('tag-group');
                const strong = document.createElement('strong');
                strong.textContent = `Base: "${escapeHtml(group.lower)}" - Variants:`;
                div.appendChild(strong);
                const ul = document.createElement('ul');
                group.originals.forEach(originalTag => {
                    const li = document.createElement('li');
                    li.textContent = escapeHtml(originalTag);
                    ul.appendChild(li);
                });
                div.appendChild(ul);
                caseInconsistentOutput.appendChild(div);
            });
            updateStatus(`Found ${inconsistentGroups.length} case-inconsistent tag group(s).`, false);
        } else {
            caseInconsistentOutput.innerHTML = '<p>No case-inconsistent tags found.</p>';
            updateStatus('No case-inconsistent tags found.', false);
        }
    }

    function exportUniqueTagList() {
        if (uniqueTags.size === 0) {
            updateStatus("No unique tags to export.");
            return;
        }
        const tagArray = [...uniqueTags].sort();
        const fileContent = tagArray.join('\n');
        const blob = new Blob([fileContent], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'unique_tags_export.txt';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        updateStatus(`Exported ${tagArray.length} unique tags.`);
    }


    async function saveAllChanges() {
        if (!datasetHandle) {
            updateStatus("No dataset loaded.");
            return;
        }
        const itemsToSave = allImageData.filter(item => item.modified);
        const deletionsOccurred = allImageData.length < initialTotalCount;
        if (itemsToSave.length === 0 && !deletionsOccurred) {
            updateStatus("No changes (modifications or deletions) to save.");
            saveAllBtn.disabled = true;
            return;
        }
        let savedCount = 0,
            errorCount = 0;
        let finalMessage = "";
        if (itemsToSave.length > 0) {
            updateStatus(`Saving ${itemsToSave.length} modified tag file(s)...`, true);
            try {
                if (await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted') {
                    if (await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted') {
                        updateStatus('Error: Write permission denied. Cannot save tags.');
                        saveAllBtn.disabled = false;
                        return;
                    }
                }
            } catch (permError) {
                updateStatus(`Error requesting permission: ${permError.message}`);
                saveAllBtn.disabled = false;
                return;
            }
            for (const item of itemsToSave) {
                const newTagString = joinTags(item.tags);
                const tagFileName = item.originalFullTagName || `${item.imageName}${TAG_EXTENSION}`; // Use existing if known, else derive
                const fullLogPath = item.relativePath ? `${item.relativePath}/${tagFileName}` : tagFileName;
                try {
                    let tagHandle = item.tagHandle;
                    if (!tagHandle) {
                        try {
                            let parentDirHandle = datasetHandle;
                            if (item.relativePath) {
                                const pathParts = item.relativePath.split('/');
                                for (const part of pathParts) {
                                    if (part) {
                                        parentDirHandle = await parentDirHandle.getDirectoryHandle(part, { create: false });
                                    }
                                }
                            }
                            tagHandle = await parentDirHandle.getFileHandle(tagFileName, { create: true });
                            item.tagHandle = tagHandle;
                            item.originalFullTagName = tagFileName; // Update if we just created it
                        } catch (handleError) {
                            console.error(`Error getting/creating tag file handle for ${fullLogPath}:`, handleError);
                            errorCount++;
                            item.modified = true;
                            continue;
                        }
                    }
                    const writable = await tagHandle.createWritable();
                    await writable.write(newTagString);
                    await writable.close();
                    item.originalTags = newTagString;
                    item.modified = false;
                    savedCount++;
                    document.getElementById(item.id)?.classList.remove('modified');
                } catch (writeError) {
                    console.error(`Error saving file ${fullLogPath}:`, writeError);
                    errorCount++;
                    item.modified = true;
                }
            }
            finalMessage = `Saved ${savedCount} tag file(s).`;
            if (errorCount > 0) {
                finalMessage += ` Failed to save ${errorCount}. Check console.`;
            }
        } else {
            finalMessage = "No tag modifications to save.";
        }
        if (deletionsOccurred) {
            finalMessage += " (Image file deletions are performed immediately).";
            initialTotalCount = allImageData.length;
        }
        const stillModified = allImageData.some(item => item.modified);
        saveAllBtn.disabled = !stillModified && !(allImageData.length < initialTotalCount);
        recalculateAllTagStats();
        updateStatus(finalMessage, false);
    }

    function navigateEditor(direction) {
        if (currentEditIndex === -1 || displayedImageDataIndices.length <= 1) return;
        let currentItemGlobalIndex = currentEditIndex;
        let currentItemDisplayIndex = -1;
        for (let i = 0; i < displayedImageDataIndices.length; i++) {
            if (displayedImageDataIndices[i] === currentItemGlobalIndex) {
                currentItemDisplayIndex = i;
                break;
            }
        }
        if (currentItemDisplayIndex === -1) {
            console.warn("Current edited item not found for nav.");
            return;
        }
        let nextItemDisplayIndex;
        if (direction === 'next') {
            nextItemDisplayIndex = (currentItemDisplayIndex + 1) % displayedImageDataIndices.length;
        } else {
            nextItemDisplayIndex = (currentItemDisplayIndex - 1 + displayedImageDataIndices.length) % displayedImageDataIndices.length;
        }
        if (editorTagsTextarea.value !== currentEditOriginalTagsString) {
            saveEditorChanges();
        }
        const nextItemGlobalIndex = displayedImageDataIndices[nextItemDisplayIndex];
        if (allImageData[nextItemGlobalIndex]) {
            openEditor(allImageData[nextItemGlobalIndex].id);
        }
    }

    function updateEditorNavButtonStates() {
        if (displayedImageDataIndices.length <= 1) {
            editorPrevBtn.disabled = true;
            editorNextBtn.disabled = true;
        } else {
            editorPrevBtn.disabled = false;
            editorNextBtn.disabled = false;
        }
    }

    function showSuggestions(inputElement) {
        suggestionInputTarget = inputElement;
        const value = inputElement.value;
        const lastCommaIndex = value.lastIndexOf(',');
        const currentTagFragment = (lastCommaIndex === -1 ? value : value.substring(lastCommaIndex + 1)).trimStart();
        if (currentTagFragment.length === 0 && !value.endsWith(',')) {
            hideSuggestions();
            return;
        }
        const existingTagsInInput = parseTags(value.substring(0, lastCommaIndex + 1));
        currentSuggestions = [...uniqueTags].filter(tag => tag.toLowerCase().includes(currentTagFragment.toLowerCase()) && !existingTagsInInput.includes(tag)).slice(0, 10);
        if (currentSuggestions.length > 0) {
            editorTagSuggestionsBox.innerHTML = '';
            const ul = document.createElement('ul');
            currentSuggestions.forEach((suggestion, index) => {
                const li = document.createElement('li');
                li.textContent = suggestion;
                li.dataset.index = index;
                li.addEventListener('click', () => {
                    selectSuggestion(suggestion);
                });
                ul.appendChild(li);
            });
            editorTagSuggestionsBox.appendChild(ul);
            editorTagSuggestionsBox.classList.remove('hidden');
            activeSuggestionIndex = -1;
        } else {
            hideSuggestions();
        }
    }

    function hideSuggestions() {
        editorTagSuggestionsBox.classList.add('hidden');
        editorTagSuggestionsBox.innerHTML = '';
        currentSuggestions = [];
        activeSuggestionIndex = -1;
        suggestionInputTarget = null;
    }

    function selectSuggestion(suggestionText) {
        if (!suggestionInputTarget) return;
        const value = suggestionInputTarget.value;
        const lastCommaIndex = value.lastIndexOf(',');
        const beforeFragment = lastCommaIndex === -1 ? '' : value.substring(0, lastCommaIndex + 1).trimEnd() + (lastCommaIndex !== -1 ? ' ' : '');
        suggestionInputTarget.value = `${beforeFragment}${suggestionText}, `;
        hideSuggestions();
        suggestionInputTarget.focus();
        suggestionInputTarget.selectionStart = suggestionInputTarget.selectionEnd = suggestionInputTarget.value.length;
    }

    function updateActiveSuggestion(newIndex) {
        const items = editorTagSuggestionsBox.querySelectorAll('li');
        if (activeSuggestionIndex !== -1 && items[activeSuggestionIndex]) {
            items[activeSuggestionIndex].classList.remove('active-suggestion');
        }
        activeSuggestionIndex = newIndex;
        if (activeSuggestionIndex !== -1 && items[activeSuggestionIndex]) {
            items[activeSuggestionIndex].classList.add('active-suggestion');
            items[activeSuggestionIndex].scrollIntoView({
                block: 'nearest'
            });
        }
    }

    function openBatchEditor() {
        if (selectedItemIds.size === 0) {
            updateStatus("No items selected for batch editing.");
            return;
        }
        batchEditorModal.classList.remove('hidden');
        batchEditorSelectedCount.textContent = selectedItemIds.size;
        batchEditorAddTagsInput.value = '';
        batchEditorRemoveTagsInput.value = '';
        batchEditorStatus.textContent = '';
        const selectedItemsData = [];
        selectedItemIds.forEach(id => {
            const index = findIndexById(id);
            if (index !== -1) selectedItemsData.push(allImageData[index]);
        });
        if (selectedItemsData.length === 0) {
            batchEditorCommonTagsDiv.innerHTML = '<p>Error: No valid selected items found.</p>';
            batchEditorPartialTagsDiv.innerHTML = '<p>Error loading data.</p>';
            return;
        }
        let commonTags = new Set(selectedItemsData[0].tags);
        const allTagsInSelection = new Map();
        selectedItemsData.forEach((item, idx) => {
            const currentItemTags = new Set(item.tags);
            if (idx > 0) {
                commonTags = new Set([...commonTags].filter(tag => currentItemTags.has(tag)));
            }
            item.tags.forEach(tag => {
                allTagsInSelection.set(tag, (allTagsInSelection.get(tag) || 0) + 1);
            });
        });
        batchEditorCommonTagsDiv.innerHTML = '';
        if (commonTags.size > 0) {
            [...commonTags].sort().forEach(tag => {
                const span = document.createElement('span');
                span.classList.add('tag');
                span.textContent = tag;
                batchEditorCommonTagsDiv.appendChild(span);
            });
        } else {
            batchEditorCommonTagsDiv.innerHTML = '<p>None</p>';
        }
        batchEditorPartialTagsDiv.innerHTML = '';
        let partialTagsFound = false;
        const sortedAllTags = [...allTagsInSelection.entries()].sort((a, b) => a[0].localeCompare(b[0]));
        sortedAllTags.forEach(([tag, count]) => {
            if (!commonTags.has(tag)) {
                partialTagsFound = true;
                const span = document.createElement('span');
                span.classList.add('tag');
                span.textContent = tag;
                const countSpan = document.createElement('span');
                countSpan.classList.add('count');
                countSpan.textContent = `(${count}/${selectedItemsData.length})`;
                span.appendChild(countSpan);
                batchEditorPartialTagsDiv.appendChild(span);
            }
        });
        if (!partialTagsFound) {
            batchEditorPartialTagsDiv.innerHTML = '<p>None (or all tags are common)</p>';
        }
    }

    function applyBatchEdits() {
        const tagsToAdd = parseTags(batchEditorAddTagsInput.value);
        const tagsToRemove = parseTags(batchEditorRemoveTagsInput.value);
        if (tagsToAdd.length === 0 && tagsToRemove.length === 0) {
            batchEditorStatus.textContent = "No tags specified to add or remove.";
            batchEditorStatus.style.color = 'var(--status-color)';
            return;
        }
        let modifiedCount = 0;
        let overallStatsChanged = false;
        selectedItemIds.forEach(id => {
            const index = findIndexById(id);
            if (index === -1) return;
            const item = allImageData[index];
            const originalItemTagsString = joinTags(item.tags);
            let currentItemTags = new Set(item.tags);
            tagsToAdd.forEach(tag => currentItemTags.add(tag));
            tagsToRemove.forEach(tag => currentItemTags.delete(tag));
            const newItemTagsArray = [...currentItemTags].sort();
            const newItemTagsString = joinTags(newItemTagsArray);
            if (newItemTagsString !== originalItemTagsString) {
                const oldTagsForStatUpdate = parseTags(originalItemTagsString);
                const addedForStatUpdate = newItemTagsArray.filter(t => !oldTagsForStatUpdate.includes(t));
                const removedForStatUpdate = oldTagsForStatUpdate.filter(t => !newItemTagsArray.includes(t));
                if (addedForStatUpdate.length > 0 || removedForStatUpdate.length > 0) {
                    updateTagStatsIncremental(addedForStatUpdate, removedForStatUpdate);
                    overallStatsChanged = true;
                }
                item.tags = newItemTagsArray;
                item.modified = true;
                modifiedCount++;
                const gridItem = document.getElementById(item.id);
                if (gridItem) {
                    gridItem.classList.add('modified');
                    const tagsDiv = gridItem.querySelector('.tags');
                    if (tagsDiv) renderTagsForItem(tagsDiv, item);
                }
            }
        });
        if (modifiedCount > 0) {
            saveAllBtn.disabled = false;
            if (overallStatsChanged) recalculateAllTagStats();
            batchEditorStatus.textContent = `Changes applied to ${modifiedCount} item(s). Remember to save all.`;
            batchEditorStatus.style.color = 'var(--success-color)';
            batchEditorAddTagsInput.value = '';
            batchEditorRemoveTagsInput.value = '';
        } else {
            batchEditorStatus.textContent = "No changes made to selected items.";
            batchEditorStatus.style.color = 'var(--status-color)';
        }
    }

    function closeBatchEditor() {
        batchEditorModal.classList.add('hidden');
        batchEditorStatus.textContent = '';
    }

    // --- Custom Formatting Rules Logic ---
    function loadCustomRules() {
        try {
            const storedRules = localStorage.getItem(CUSTOM_RULES_STORAGE_KEY);
            if (storedRules) {
                customRules = JSON.parse(storedRules);
                if (!Array.isArray(customRules)) customRules = []; 
            } else {
                customRules = [];
            }
        } catch (e) {
            console.error("Error loading custom rules:", e);
            customRules = [];
            updateStatus("Error loading custom rules from storage. Check console.");
        }

        if (customRules.length === 0) {
            customRules.push({
                id: generateRuleId('default-replace-underscores'),
                name: "Default: Replace Underscores with Spaces",
                find: "_",
                replace: " ",
                isRegex: false,
                isCaseSensitive: false, 
                applyTo: "all",
                specificTags: []
            });
            customRules.push({
                id: generateRuleId('default-escape-parentheses'),
                name: "Default: Escape Parentheses () -> \\(\\) ",
                find: "(\\(|\\))", 
                replace: "\\\\$1", 
                isRegex: true,
                isCaseSensitive: true, 
                applyTo: "all",
                specificTags: []
            });
            customRules.push({
                id: generateRuleId('default-smart-brackets'),
                name: "Default: Format 'word_(info)' to 'word from info'",
                find: "^([^\\s_()]+)(?:_)?\\(([^)]+)\\)$",
                replace: "$1 from $2",
                isRegex: true,
                isCaseSensitive: false, 
                applyTo: "all",
                specificTags: []
            });
            saveCustomRules(); 
        }

        renderCustomRulesList();
        populateCustomRulesBulkApplyDropdown();
    }

    function saveCustomRules() {
        try {
            localStorage.setItem(CUSTOM_RULES_STORAGE_KEY, JSON.stringify(customRules));
        } catch (e) {
            console.error("Error saving custom rules:", e);
            updateStatus("Error saving custom rules to storage. Check console.");
        }
    }

    function renderCustomRulesList() {
        customRulesListDiv.innerHTML = '';
        if (customRules.length === 0) {
            customRulesListDiv.innerHTML = '<p>No custom rules defined yet.</p>';
            return;
        }
        const ul = document.createElement('ul');
        ul.classList.add('custom-rule-item-list');
        customRules.forEach(rule => {
            const li = document.createElement('li');
            li.dataset.ruleId = rule.id;
            const nameSpan = document.createElement('span');
            nameSpan.classList.add('rule-name');
            nameSpan.textContent = rule.name;
            li.appendChild(nameSpan);
            const summarySpan = document.createElement('span');
            summarySpan.classList.add('rule-summary');
            let findText = rule.find.length > 20 ? rule.find.substring(0, 17) + "..." : rule.find;
            let replaceText = rule.replace.length > 20 ? rule.replace.substring(0, 17) + "..." : rule.replace;
            summarySpan.textContent = `Find: "${escapeHtml(findText)}", Replace: "${escapeHtml(replaceText)}" ${rule.isRegex ? '(Regex)' : ''}`;
            li.appendChild(summarySpan);
            const actionsDiv = document.createElement('div');
            actionsDiv.classList.add('rule-actions');
            const editBtn = document.createElement('button');
            editBtn.textContent = 'Edit';
            editBtn.classList.add('edit-rule-btn');
            editBtn.addEventListener('click', () => openRuleEditor(rule.id));
            actionsDiv.appendChild(editBtn);
            const deleteBtn = document.createElement('button');
            deleteBtn.textContent = 'Delete';
            deleteBtn.classList.add('delete-rule-btn');
            deleteBtn.addEventListener('click', () => deleteCustomRule(rule.id));
            actionsDiv.appendChild(deleteBtn);
            li.appendChild(actionsDiv);
            ul.appendChild(li);
        });
        customRulesListDiv.appendChild(ul);
    }

    function populateCustomRulesBulkApplyDropdown() {
        bulkApplyCustomRulesSelect.innerHTML = '';
        if (customRules.length === 0) {
            const option = document.createElement('option');
            option.value = "";
            option.textContent = "No custom rules defined";
            option.disabled = true;
            bulkApplyCustomRulesSelect.appendChild(option);
            bulkApplyCustomRulesBtn.disabled = true;
        } else {
            customRules.forEach(rule => {
                const option = document.createElement('option');
                option.value = rule.id;
                option.textContent = rule.name;
                bulkApplyCustomRulesSelect.appendChild(option);
            });
            bulkApplyCustomRulesBtn.disabled = false;
        }
    }

    function openRuleEditor(ruleId = null) {
        editingRuleId = ruleId;
        ruleEditorModal.classList.remove('hidden');
        ruleEditorStatus.textContent = '';
        ruleSpecificTagsGroup.style.display = 'none';
        if (ruleId) {
            const rule = customRules.find(r => r.id === ruleId);
            if (!rule) {
                ruleEditorStatus.textContent = 'Error: Rule not found.';
                ruleEditorStatus.style.color = 'var(--error-color)';
                return;
            }
            ruleEditorTitle.textContent = 'Edit Custom Rule';
            ruleIdInput.value = rule.id;
            ruleNameInput.value = rule.name;
            ruleFindInput.value = rule.find;
            ruleReplaceInput.value = rule.replace;
            ruleIsRegexCheckbox.checked = rule.isRegex;
            ruleIsCaseSensitiveCheckbox.checked = rule.isRegex ? rule.isCaseSensitive : false;
            ruleIsCaseSensitiveCheckbox.disabled = !rule.isRegex;
            ruleApplyToSelect.value = rule.applyTo;
            if (rule.applyTo === 'only_if_is' || rule.applyTo === 'not_if_is') {
                ruleSpecificTagsGroup.style.display = 'block';
                ruleSpecificTagsInput.value = rule.specificTags ? rule.specificTags.join(', ') : '';
            }
        } else {
            ruleEditorTitle.textContent = 'Create New Custom Rule';
            ruleIdInput.value = generateRuleId();
            ruleNameInput.value = '';
            ruleFindInput.value = '';
            ruleReplaceInput.value = '';
            ruleIsRegexCheckbox.checked = false;
            ruleIsCaseSensitiveCheckbox.checked = false;
            ruleIsCaseSensitiveCheckbox.disabled = true;
            ruleApplyToSelect.value = 'all';
            ruleSpecificTagsInput.value = '';
        }
        ruleNameInput.focus();
    }

    function closeRuleEditor() {
        ruleEditorModal.classList.add('hidden');
        editingRuleId = null;
    }
    ruleIsRegexCheckbox.addEventListener('change', () => {
        ruleIsCaseSensitiveCheckbox.disabled = !ruleIsRegexCheckbox.checked;
        if (!ruleIsRegexCheckbox.checked) {
            ruleIsCaseSensitiveCheckbox.checked = false;
        }
    });
    ruleApplyToSelect.addEventListener('change', () => {
        if (ruleApplyToSelect.value === 'only_if_is' || ruleApplyToSelect.value === 'not_if_is') {
            ruleSpecificTagsGroup.style.display = 'block';
        } else {
            ruleSpecificTagsGroup.style.display = 'none';
        }
    });
    ruleEditorSaveBtn.addEventListener('click', () => {
        const id = ruleIdInput.value;
        const name = ruleNameInput.value.trim();
        const findPattern = ruleFindInput.value;
        const replacePattern = ruleReplaceInput.value;
        const isRegex = ruleIsRegexCheckbox.checked;
        const isCaseSensitive = ruleIsCaseSensitiveCheckbox.checked;
        const applyTo = ruleApplyToSelect.value;
        const specificTags = parseTags(ruleSpecificTagsInput.value);
        if (!name) {
            ruleEditorStatus.textContent = 'Rule name is required.';
            ruleEditorStatus.style.color = 'var(--error-color)';
            return;
        }
        if (!findPattern && !isRegex) { // findPattern can be empty for regex (e.g. ^ to prepend)
            ruleEditorStatus.textContent = 'Find pattern is required for non-regex rules.';
            ruleEditorStatus.style.color = 'var(--error-color)';
            return;
        }
        if ((applyTo === 'only_if_is' || applyTo === 'not_if_is') && specificTags.length === 0) {
            ruleEditorStatus.textContent = 'Please provide specific tags for this "Apply To" condition.';
            ruleEditorStatus.style.color = 'var(--error-color)';
            return;
        }
        if (isRegex) {
            try {
                new RegExp(findPattern);
            } catch (e) {
                ruleEditorStatus.textContent = `Invalid Regular Expression in Find pattern: ${e.message}`;
                ruleEditorStatus.style.color = 'var(--error-color)';
                return;
            }
        }
        const ruleData = {
            id,
            name,
            find: findPattern,
            replace: replacePattern,
            isRegex,
            isCaseSensitive,
            applyTo,
            specificTags
        };
        if (editingRuleId) {
            const index = customRules.findIndex(r => r.id === editingRuleId);
            if (index !== -1) {
                customRules[index] = ruleData;
            } else {
                customRules.push(ruleData);
            }
        } else {
            customRules.push(ruleData);
        }
        saveCustomRules();
        renderCustomRulesList();
        populateCustomRulesBulkApplyDropdown();
        closeRuleEditor();
        updateStatus(`Custom rule "${name}" ${editingRuleId ? 'updated' : 'created'}.`);
    });

    function deleteCustomRule(ruleId) {
        if (confirm(`Are you sure you want to delete the custom rule: "${customRules.find(r=>r.id===ruleId)?.name || 'this rule'}"?`)) {
            customRules = customRules.filter(r => r.id !== ruleId);
            saveCustomRules();
            renderCustomRulesList();
            populateCustomRulesBulkApplyDropdown();
            updateStatus('Custom rule deleted.');
        }
    }

    function applySingleCustomRuleToTag(tag, rule) {
        let newTag = tag;
        const applyThisRule = () => {
            if (rule.isRegex) {
                try {
                    const flags = 'g' + (rule.isCaseSensitive ? '' : 'i');
                    const regex = new RegExp(rule.find, flags);
                    newTag = newTag.replace(regex, rule.replace);
                } catch (e) {
                    console.warn(`Error applying regex rule "${rule.name}" to tag "${tag}": ${e.message}`);
                    return tag;
                }
            } else {
                if (rule.find === '') return newTag; 
                if (rule.isCaseSensitive) { // Stricter interpretation: non-regex case sensitive means direct string replacement
                    newTag = newTag.split(rule.find).join(rule.replace);
                } else { // Non-regex, case-insensitive
                    const escFind = rule.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                    const regex = new RegExp(escFind, 'gi');
                    newTag = newTag.replace(regex, rule.replace);
                }
            }
            return newTag;
        };
        switch (rule.applyTo) {
            case 'all':
                return applyThisRule();
            case 'only_if_is':
                if (rule.specificTags.includes(tag)) {
                    return applyThisRule();
                }
                break;
            case 'not_if_is':
                if (!rule.specificTags.includes(tag)) {
                    return applyThisRule();
                }
                break;
        }
        return tag;
    }

    function performBulkCustomRulesOperation() {
        const selectedRuleIds = Array.from(bulkApplyCustomRulesSelect.selectedOptions).map(opt => opt.value);
        if (selectedRuleIds.length === 0) {
            updateStatus("No custom rules selected to apply.");
            return;
        }
        const rulesToApply = selectedRuleIds.map(id => customRules.find(r => r.id === id)).filter(r => r);
        if (rulesToApply.length === 0) {
            updateStatus("Selected custom rules not found. This should not happen.");
            return;
        }
        if (!confirm(`Apply ${rulesToApply.length} selected custom rule(s) to ALL images?`)) {
            updateStatus("Custom rule application cancelled.");
            return;
        }
        updateStatus(`Applying ${rulesToApply.length} custom rule(s)...`, true);
        let modCount = 0;
        let affectedItemIds = new Set();
        let overallStatsChanged = false;
        allImageData.forEach(item => {
            let originalItemTagsString = joinTags(item.tags);
            let currentItemTagsArray = [...item.tags];
            rulesToApply.forEach(rule => {
                currentItemTagsArray = currentItemTagsArray.map(tag => applySingleCustomRuleToTag(tag, rule));
            });
            currentItemTagsArray = [...new Set(currentItemTagsArray.map(t => t.trim()).filter(t => t))]; 
            if (joinTags(currentItemTagsArray) !== originalItemTagsString) {
                item.tags = currentItemTagsArray;
                item.modified = true;
                modCount++;
                affectedItemIds.add(item.id);
                overallStatsChanged = true;
            }
        });
        if (modCount > 0) {
            saveAllBtn.disabled = false;
            affectedItemIds.forEach(id => {
                const gridItem = document.getElementById(id);
                const itemData = allImageData.find(i => i.id === id);
                if (gridItem && itemData) {
                    gridItem.classList.add('modified');
                    const tagsDiv = gridItem.querySelector('.tags');
                    if (tagsDiv) renderTagsForItem(tagsDiv, itemData);
                }
            });
            if (overallStatsChanged) recalculateAllTagStats();
            updateStatus(`${modCount} items affected by custom rules. Save changes.`, false);
        } else {
            updateStatus(`Custom rules applied. No changes made to any items.`, false);
        }
    }

    // --- Global Event Handlers ---
    function handleGlobalKeyDown(event) {
        const activeElement = document.activeElement;
        const isInputFocusedGeneral = ['INPUT', 'TEXTAREA', 'SELECT'].includes(activeElement?.tagName);
        const isSingleEditorOpen = !editorModal.classList.contains('hidden');
        const isBatchEditorOpen = !batchEditorModal.classList.contains('hidden');
        const isRuleEditorOpen = !ruleEditorModal.classList.contains('hidden');
        const isModalOpen = isSingleEditorOpen || isBatchEditorOpen || isRuleEditorOpen;
        const isTextareaFocused = editorTagsTextarea.isSameNode(activeElement);
        const isImageZoomed = editorImagePreview.classList.contains('zoomed');
        const suggestionsVisible = !editorTagSuggestionsBox.classList.contains('hidden');
        if ((event.ctrlKey || event.metaKey) && event.key === 's') {
            event.preventDefault();
            if (!saveAllBtn.disabled) saveAllChanges();
            return;
        }
        if (isTextareaFocused && suggestionsVisible) {
            if (event.key === 'ArrowDown') {
                event.preventDefault();
                updateActiveSuggestion((activeSuggestionIndex + 1) % currentSuggestions.length);
                return;
            }
            if (event.key === 'ArrowUp') {
                event.preventDefault();
                updateActiveSuggestion((activeSuggestionIndex - 1 + currentSuggestions.length) % currentSuggestions.length);
                return;
            }
            if ((event.key === 'Enter' || event.key === 'Tab') && activeSuggestionIndex !== -1) {
                event.preventDefault();
                selectSuggestion(currentSuggestions[activeSuggestionIndex]);
                return;
            }
            if (event.key === 'Escape') {
                event.preventDefault();
                hideSuggestions();
                return;
            }
        }
        if (event.key === 'Escape') {
            if (isImageZoomed) {
                toggleEditorImageZoom();
                return;
            }
            if (isRuleEditorOpen) {
                closeRuleEditor();
                return;
            }
            if (isSingleEditorOpen) closeEditor();
            else if (isBatchEditorOpen) closeBatchEditor();
            else if (selectedItemIds.size > 0) clearSelection();
            else if (isInputFocusedGeneral && activeElement.value !== '') {
                activeElement.value = '';
                if (activeElement.id === 'tag-frequency-search') activeElement.dispatchEvent(new Event('input', {
                    bubbles: true
                }));
            } else if (keyboardFocusIndex !== -1) clearKeyboardFocus();
            return;
        }
        if (event.key === 'Enter' && isInputFocusedGeneral && !isModalOpen) {
            let buttonToClick = null;
            switch (activeElement.id) {
                case 'search-tags':
                case 'exclude-tags':
                    buttonToClick = filterBtn;
                    break;
                case 'rename-old-tag':
                case 'rename-new-tag':
                    buttonToClick = renameTagBtn;
                    break;
                case 'remove-tag-name':
                    buttonToClick = removeTagBtn;
                    break;
                case 'add-trigger-words':
                    buttonToClick = addTagsBtn;
                    break;
                case 'tag-frequency-search':
                    event.preventDefault();
                    return;
                case 'sort-property':
                case 'sort-order':
                    buttonToClick = applySortBtn;
                    break;
            }
            if (buttonToClick && !buttonToClick.disabled) {
                event.preventDefault();
                buttonToClick.click();
                return;
            }
        }
        if (isSingleEditorOpen && !isTextareaFocused && !isImageZoomed && (event.key === 'ArrowLeft' || event.key === 'ArrowRight')) {
            event.preventDefault();
            navigateEditor(event.key === 'ArrowLeft' ? 'previous' : 'next');
            return;
        }
        if (!isModalOpen && !isInputFocusedGeneral && displayedImageDataIndices.length > 0) {
            let newIndex = keyboardFocusIndex;
            let handled = false;
            if (newIndex === -1 && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
                newIndex = 0;
                handled = true;
            } else {
                let columns = 1;
                const gridWidth = imageGrid.offsetWidth;
                const firstItemElement = imageGrid.querySelector('.grid-item');
                if (firstItemElement) {
                    const itemStyle = getComputedStyle(firstItemElement);
                    const itemOuterWidth = firstItemElement.offsetWidth + parseInt(itemStyle.marginLeft) + parseInt(itemStyle.marginRight);
                    if (itemOuterWidth > 0) columns = Math.max(1, Math.floor(gridWidth / itemOuterWidth));
                }
                switch (event.key) {
                    case 'ArrowLeft':
                        newIndex = Math.max(0, newIndex - 1);
                        handled = true;
                        break;
                    case 'ArrowRight':
                        newIndex = Math.min(displayedImageDataIndices.length - 1, newIndex + 1);
                        handled = true;
                        break;
                    case 'ArrowDown':
                        newIndex = Math.min(displayedImageDataIndices.length - 1, newIndex + columns);
                        handled = true;
                        break;
                    case 'ArrowUp':
                        newIndex = Math.max(0, newIndex - columns);
                        handled = true;
                        break;
                    case 'Enter':
                        if (keyboardFocusIndex !== -1) {
                            const globalIdx = displayedImageDataIndices[keyboardFocusIndex];
                            const itemId = allImageData[globalIdx]?.id;
                            if (itemId) {
                                clearSelection();
                                addSelection(itemId);
                                openEditor(itemId);
                            }
                            handled = true;
                        }
                        break;
                }
            }
            if (handled) {
                event.preventDefault();
                if (newIndex !== keyboardFocusIndex || event.key === 'Enter') setKeyboardFocus(newIndex);
            }
        }
    }

    function toggleEditorImageZoom() {
        editorImagePreview.classList.toggle('zoomed');
        editorModalContent.classList.toggle('image-zoomed');
    }

    function handleDocumentClickToClearSelection(event) {
        const isModalOpen = !editorModal.classList.contains('hidden') || !batchEditorModal.classList.contains('hidden') || !ruleEditorModal.classList.contains('hidden');
        if (isModalOpen) return;
        const clickedOnGridItem = event.target.closest('.grid-item');
        const clickedOnBatchEditBtn = event.target.closest('#batch-edit-selection-btn');
        if (!clickedOnGridItem && !clickedOnBatchEditBtn) {
            if (selectedItemIds.size > 0) clearSelection();
        }
    }

    // --- Event Listeners Setup ---
    loadDatasetBtn.addEventListener('click', loadDataset);
    saveAllBtn.addEventListener('click', saveAllChanges);
    themeToggleBtn.addEventListener('click', toggleTheme);
    batchEditSelectionBtn.addEventListener('click', openBatchEditor);
    filterBtn.addEventListener('click', filterAndSearch);
    resetFilterBtn.addEventListener('click', resetFilters);
    applySortBtn.addEventListener('click', () => {
        currentSortProperty = sortPropertySelect.value;
        currentSortOrder = sortOrderSelect.value;
        filterAndSearch(); 
    });
    renameTagBtn.addEventListener('click', () => performStandardBulkOperation('rename'));
    removeTagBtn.addEventListener('click', () => performStandardBulkOperation('remove'));
    addTagsBtn.addEventListener('click', () => performStandardBulkOperation('add'));
    removeDuplicatesBtn.addEventListener('click', () => performStandardBulkOperation('removeDuplicates'));
    
    // Utility Button Listeners
    findDuplicateFilesBtn.addEventListener('click', findDuplicateTagFiles);
    findOrphanedMissingBtn.addEventListener('click', findOrphanedAndMissingTagFiles);
    findCaseInconsistentTagsBtn.addEventListener('click', findCaseInconsistentTags);
    exportUniqueTagsBtn.addEventListener('click', exportUniqueTagList);

    createNewRuleBtn.addEventListener('click', () => openRuleEditor());
    ruleEditorCancelBtn.addEventListener('click', closeRuleEditor);
    ruleEditorModal.addEventListener('click', (e) => {
        if (e.target === ruleEditorModal) closeRuleEditor();
    });
    bulkApplyCustomRulesBtn.addEventListener('click', performBulkCustomRulesOperation);
    tagFrequencySearchInput.addEventListener('input', () => {
        clearTimeout(debouncedUpdateFrequencyList);
        debouncedUpdateFrequencyList = setTimeout(updateStatisticsDisplay, 300);
    });
    editorSaveBtn.addEventListener('click', () => {
        saveEditorChanges();
        hideSuggestions();
    });
    editorCancelBtn.addEventListener('click', closeEditor);
    editorRemoveDupsBtn.addEventListener('click', () => {
        removeDuplicateTagsInEditor();
        hideSuggestions();
    });
    editorModal.addEventListener('click', (e) => {
        if (e.target === editorModal && !editorImagePreview.classList.contains('zoomed')) closeEditor();
    });
    batchEditorApplyBtn.addEventListener('click', applyBatchEdits);
    batchEditorCancelBtn.addEventListener('click', closeBatchEditor);
    batchEditorModal.addEventListener('click', (e) => {
        if (e.target === batchEditorModal) closeBatchEditor();
    });
    editorTagsTextarea.addEventListener('input', (e) => {
        showSuggestions(e.target);
    });
    editorTagsTextarea.addEventListener('blur', () => {
        setTimeout(() => {
            if (!editorTagSuggestionsBox.contains(document.activeElement)) hideSuggestions();
        }, 150);
    });
    editorTagsTextarea.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
            e.preventDefault();
            saveEditorChanges();
            hideSuggestions();
        }
    });
    editorImagePreview.addEventListener('click', toggleEditorImageZoom);
    if (editorPrevBtn) editorPrevBtn.addEventListener('click', () => navigateEditor('previous'));
    if (editorNextBtn) editorNextBtn.addEventListener('click', () => navigateEditor('next'));
    document.addEventListener('keydown', handleGlobalKeyDown);
    document.addEventListener('click', handleDocumentClickToClearSelection);

    // --- Initial Setup ---
    loadInitialTheme();
    loadCustomRules();
    sortPropertySelect.value = currentSortProperty;
    sortOrderSelect.value = currentSortOrder;
    updateStatus('Ready. Load a dataset folder.');
    if (!window.showDirectoryPicker) {
        updateStatus("Warning: Browser lacks File System Access API support. Core functionality will be limited.");
        loadDatasetBtn.disabled = true;
        saveAllBtn.disabled = true;
        batchEditSelectionBtn.disabled = true;
        applySortBtn.disabled = true;
        createNewRuleBtn.disabled = true;
        bulkApplyCustomRulesBtn.disabled = true;
        [renameTagBtn, removeTagBtn, addTagsBtn, removeDuplicatesBtn, 
         findDuplicateFilesBtn, findOrphanedMissingBtn, findCaseInconsistentTagsBtn, exportUniqueTagsBtn]
        .forEach(btn => {
            if (btn) btn.disabled = true;
        });
    }
});