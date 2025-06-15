document.addEventListener('DOMContentLoaded', () => {
    // --- DOM Elements ---
    const loadDatasetBtn = document.getElementById('load-dataset-btn');
    const saveAllBtn = document.getElementById('save-all-btn');
    const emptyTrashBtn = document.getElementById('empty-trash-btn');
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

    // Unification DOM Elements
    const unifyPrefixInput = document.getElementById('unify-prefix');
    const unifyStartNumInput = document.getElementById('unify-start-number');
    const unifyPadZerosCheckbox = document.getElementById('unify-pad-zeros');
    const unifyPaddingLengthGroup = document.getElementById('unify-padding-length-group');
    const unifyPaddingLengthInput = document.getElementById('unify-padding-length');
    const unifyPreviewBtn = document.getElementById('unify-preview-btn');
    const unifyPreviewOutput = document.getElementById('unify-preview-output');
    const unifyApplyBtn = document.getElementById('unify-apply-btn');
    const unifyStatus = document.getElementById('unify-status');

    // Image Format Conversion DOM Elements
    const convertFormatSelect = document.getElementById('convert-format-select');
    const convertQualityGroup = document.getElementById('convert-quality-group');
    const convertQualitySlider = document.getElementById('convert-quality-slider');
    const convertQualityValue = document.getElementById('convert-quality-value');
    const convertPreviewBtn = document.getElementById('convert-preview-btn');
    const convertPreviewOutput = document.getElementById('convert-preview-output');
    const convertApplyBtn = document.getElementById('convert-apply-btn');
    const convertStatus = document.getElementById('convert-status');


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
    const batchEditorMoveToTrashBtn = document.getElementById('batch-editor-move-to-trash-btn');
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
    // Holds { imageHandle, tagHandle, imageName, relativePath, imageUrl, tags, originalTags, modified, id, fileSize, lastModified, imageWidth, imageHeight, isTrashed }
    let allImageData = []; 
    let allFileSystemEntries = new Map(); // Stores { handle, kind, name } for all scanned entries, keyed by full relative path
    let displayedImageDataIndices = []; // Indices into allImageData
    let currentEditIndex = -1; // Index in allImageData
    let currentEditOriginalTagsString = '';
    let tagFrequencies = new Map();
    let uniqueTags = new Set();
    let nextItemId = 0;
    let debouncedUpdateFrequencyList = null;
    let keyboardFocusIndex = -1; // Index in displayedImageDataIndices

    let selectedItemIds = new Set();
    let lastInteractedItemId = null;

    let currentSuggestions = [];
    let activeSuggestionIndex = -1;
    let suggestionInputTarget = null;

    let currentSortProperty = 'path';
    let currentSortOrder = 'asc';

    let customRules = [];
    let editingRuleId = null;

    let unificationPreviewData = []; // Stores { itemId, oldImageName, newImageName, oldTagName, newTagName, relativePath }
    let conversionPreviewData = []; // Stores { itemId, oldImageName, newImageName, oldTagName, newTagName, relativePath, targetFormat, quality }


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
        if (typeof unsafe !== 'string') return '';
        return unsafe
             .replace(/&/g, "&amp;")
             .replace(/</g, "&lt;")
             .replace(/>/g, "&gt;")
             .replace(/"/g, "&quot;")
             .replace(/'/g, "&#039;");
    }

    function getActiveItems() {
        return allImageData.filter(item => !item.isTrashed);
    }
    
    function getActiveItemById(itemId) {
        const item = allImageData.find(i => i.id === itemId);
        return (item && !item.isTrashed) ? item : null;
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
                    
                    let pair = fileHandlesMap.get(mapKey) || {
                        imageHandle: null,
                        tagHandle: null,
                        originalImageName: null, 
                        originalTagName: null,   
                        baseNameKey: mapKey,     
                        relativePath: currentPath
                    };

                    if (IMAGE_EXTENSIONS.includes(ext)) {
                        pair.imageHandle = entry;
                        pair.originalImageName = entry.name; 
                    } else if (ext === TAG_EXTENSION || entry.name.endsWith(TAG_EXTENSION)) { // Also catch .png.txt etc
                        pair.tagHandle = entry;
                        pair.originalTagName = entry.name;
                    }
                    
                    if (IMAGE_EXTENSIONS.includes(ext) || ext === TAG_EXTENSION || entry.name.endsWith(TAG_EXTENSION)) {
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
            let fileHandlesMap = new Map(); 
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
                if (pair.imageHandle) { 
                    let tags = [];
                    let originalTags = '';
                    let tagHandle = pair.tagHandle;
                    
                    // Ensure imageName is just the base, without any extension.
                    const lastDotIndex = pair.originalImageName.lastIndexOf('.');
                    const imageNameForRecord = lastDotIndex > -1 ? pair.originalImageName.substring(0, lastDotIndex) : pair.originalImageName;


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
                            imageName: imageNameForRecord, // Base name without extension
                            originalFullImageName: pair.originalImageName, // Full name with original extension
                            originalFullTagName: pair.originalTagName, // Full original tag name, if exists
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
                            isTrashed: false, 
                        });
                        updateTagStats(tags, [], true); 
                    } catch (e) {
                        console.warn(`Image processing error for ${imageNameForRecord}: ${e.message}`);
                    }
                }
                processedCount++;
                if (progressElement) progressElement.value = processedCount;
            }

            displayedImageDataIndices = getActiveItems().map(item => allImageData.indexOf(item));
            currentSortProperty = sortPropertySelect.value;
            currentSortOrder = sortOrderSelect.value;
            sortAndRender(); 
            updateSaveAllButtonState();
            updateEmptyTrashButtonState();
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
        updateSaveAllButtonState();
        updateEmptyTrashButtonState();
        duplicateFilesOutput.innerHTML = '';
        orphanedMissingOutput.innerHTML = '';
        caseInconsistentOutput.innerHTML = '';
        unifyPreviewOutput.innerHTML = ''; unifyApplyBtn.disabled = true; unifyStatus.textContent = ''; unificationPreviewData = [];
        convertPreviewOutput.innerHTML = ''; convertApplyBtn.disabled = true; convertStatus.textContent = ''; conversionPreviewData = [];
        currentEditIndex = -1;
        nextItemId = 0;
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
        if (keyboardFocusIndex !== -1 && displayedImageDataIndices.length > 0 && displayedImageDataIndices[keyboardFocusIndex] !== undefined ) {
            const globalIndex = displayedImageDataIndices[keyboardFocusIndex];
            if (allImageData[globalIndex]) {
                 const oldItemId = allImageData[globalIndex].id;
                 document.getElementById(oldItemId)?.classList.remove('keyboard-focus');
            }
        }
        keyboardFocusIndex = -1;
    }

    function setKeyboardFocus(newIndex) { // newIndex is index into displayedImageDataIndices
        if (newIndex < 0 || newIndex >= displayedImageDataIndices.length) return;
        clearKeyboardFocus();
        keyboardFocusIndex = newIndex;
        const globalIndex = displayedImageDataIndices[keyboardFocusIndex];
        if (allImageData[globalIndex]) {
            const currentItemId = allImageData[globalIndex].id;
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
            if (item.isTrashed) div.classList.add('trashed-item');
            if (selectedItemIds.has(item.id)) div.classList.add('selected');

            div.tabIndex = -1; 
            div.style.outline = 'none'; 

            div.addEventListener('click', (e) => handleGridItemInteraction(e, item.id, globalIndex, displayIndex));

            const actionBtn = document.createElement('span');
            actionBtn.classList.add('image-action-btn');
            if (item.isTrashed) {
                actionBtn.classList.add('restore-btn');
                actionBtn.textContent = '♻'; 
                actionBtn.title = `Restore "${item.originalFullImageName}" from trash`;
                actionBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    restoreFromTrash(item.id);
                });
            } else {
                actionBtn.classList.add('move-to-trash-btn');
                actionBtn.textContent = '×'; 
                const deleteTitle = item.relativePath ? `Move "${item.relativePath}/${item.originalFullImageName}" to trash` : `Move "${item.originalFullImageName}" to trash`;
                actionBtn.title = deleteTitle;
                actionBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    promptMoveToTrash(item.id);
                });
            }
            div.appendChild(actionBtn);

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
        const item = allImageData[globalItemIndex];
        if (!item) return; 

        const isImageClick = event.target.tagName === 'IMG';
        const canOpenEditor = !item.isTrashed && isImageClick;

        if (event.shiftKey && lastInteractedItemId && lastInteractedItemId !== itemId) {
            const lastInteractedGlobalIndex = findIndexById(lastInteractedItemId);
            const lastInteractedDisplayIndex = displayedImageDataIndices.indexOf(lastInteractedGlobalIndex);
            const currentDisplayIndex = displayItemIndex;

            if (lastInteractedDisplayIndex === -1 || currentDisplayIndex === -1) { 
                const wasSelected = selectedItemIds.has(itemId);
                const wasOnlySelection = wasSelected && selectedItemIds.size === 1;
                if(wasOnlySelection) {
                    selectedItemIds.delete(itemId);
                    document.getElementById(itemId)?.classList.remove('selected');
                    lastInteractedItemId = null;
                } else {
                    clearSelection();
                    if (!item.isTrashed) addSelection(itemId); 
                    lastInteractedItemId = item.isTrashed ? null : itemId;
                }
            } else {
                const start = Math.min(lastInteractedDisplayIndex, currentDisplayIndex);
                const end = Math.max(lastInteractedDisplayIndex, currentDisplayIndex);

                if (!(event.ctrlKey || event.metaKey)) { 
                    clearSelection();
                }
                for (let i = start; i <= end; i++) {
                    const gIdx = displayedImageDataIndices[i];
                    if (allImageData[gIdx] && !allImageData[gIdx].isTrashed) { 
                        addSelection(allImageData[gIdx].id);
                    }
                }
            }
        } else if (event.ctrlKey || event.metaKey) {
            if (!item.isTrashed) toggleSelection(itemId); 
            lastInteractedItemId = item.isTrashed ? null : itemId;
        } else { 
            const wasSelected = selectedItemIds.has(itemId);
            const wasOnlySelection = wasSelected && selectedItemIds.size === 1;

            if (wasOnlySelection) { 
                selectedItemIds.delete(itemId);
                document.getElementById(itemId)?.classList.remove('selected');
                lastInteractedItemId = null;
            } else {
                clearSelection();
                if (!item.isTrashed) {
                    addSelection(itemId);
                    lastInteractedItemId = itemId;
                    if (canOpenEditor) { 
                         openEditor(itemId);
                    }
                } else {
                    lastInteractedItemId = null; 
                }
            }
        }
        updateBatchEditButtonState();
    }

    function toggleSelection(itemId) {
        const element = document.getElementById(itemId);
        const item = allImageData[findIndexById(itemId)];
        if (item && item.isTrashed) return; 

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
        const item = allImageData[findIndexById(itemId)];
        if (item && item.isTrashed) return; 

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

                if (!item.isTrashed) { 
                    const deleteBtn = document.createElement('span');
                    deleteBtn.classList.add('tag-delete-btn');
                    deleteBtn.textContent = '×';
                    deleteBtn.title = `Remove tag "${tag}"`;
                    deleteBtn.addEventListener('click', (e) => {
                        e.stopPropagation(); 
                        removeTagFromImage(item.id, tag);
                    });
                    tagSpan.appendChild(deleteBtn);
                }
                tagsContainer.appendChild(tagSpan);
            });
        }
    }

    function updateDisplayedCount() {
        const activeItemsCount = getActiveItems().length;
        displayedCountSpan.textContent = displayedImageDataIndices.length; 
        totalCountSpan.textContent = activeItemsCount; 
    }

    function recalculateAllTagStats() { 
        tagFrequencies.clear();
        uniqueTags.clear();
        getActiveItems().forEach(item => {
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
        const activeItems = getActiveItems();
        statsTotalImages.textContent = activeItems.length;
        statsUniqueTags.textContent = uniqueTags.size;
        statsModifiedImages.textContent = activeItems.filter(item => item.modified).length;

        tagFrequencyList.innerHTML = '';
        const searchTerm = tagFrequencySearchInput.value.trim().toLowerCase();
        const filteredFrequencies = [...tagFrequencies.entries()].filter(([t]) => !searchTerm || t.toLowerCase().includes(searchTerm));
        const sortedTags = filteredFrequencies.sort((a, b) => (b[1] - a[1]) || a[0].localeCompare(b[0]));

        if (sortedTags.length === 0) {
            tagFrequencyList.innerHTML = `<div class="tag-frequency-item"><span>${searchTerm ? 'No active tags match search.' : 'No active tags loaded.'}</span></div>`;
        } else {
            sortedTags.forEach(([tag, count]) => {
                const div = document.createElement('div');
                div.classList.add('tag-frequency-item');

                const tagInfoDiv = document.createElement('div');
                tagInfoDiv.classList.add('tag-info');

                const spanT = document.createElement('span');
                spanT.classList.add('tag-name');
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
                tagInfoDiv.appendChild(spanT);

                const spanC = document.createElement('span');
                spanC.classList.add('tag-count');
                spanC.textContent = count;
                tagInfoDiv.appendChild(spanC);
                div.appendChild(tagInfoDiv);

                const actionsDiv = document.createElement('div');
                actionsDiv.classList.add('tag-actions');

                const deleteBtn = document.createElement('button');
                deleteBtn.textContent = 'Del';
                deleteBtn.title = `Delete all occurrences of "${tag}" from active images`;
                deleteBtn.classList.add('small-action-btn', 'danger-btn');
                deleteBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    handleDeleteTagFromFrequencyList(tag);
                });
                actionsDiv.appendChild(deleteBtn);

                const renameBtn = document.createElement('button');
                renameBtn.textContent = 'Ren';
                renameBtn.title = `Rename all occurrences of "${tag}" in active images`;
                renameBtn.classList.add('small-action-btn', 'warning-btn');
                renameBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    handleRenameTagFromFrequencyList(tag);
                });
                actionsDiv.appendChild(renameBtn);
                div.appendChild(actionsDiv);

                tagFrequencyList.appendChild(div);
            });
        }
        updateDisplayedCount();
    }

    function updateSaveAllButtonState() {
        const hasModifiedActiveItems = getActiveItems().some(item => item.modified);
        saveAllBtn.disabled = !hasModifiedActiveItems;
    }

    function updateEmptyTrashButtonState() {
        const trashedCount = allImageData.filter(item => item.isTrashed).length;
        emptyTrashBtn.textContent = `Empty Trash (${trashedCount})`;
        emptyTrashBtn.disabled = trashedCount === 0;
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
            let baseItemsPoolIndices;
            if (filterMode === 'trashed') {
                baseItemsPoolIndices = allImageData.map((item, i) => item.isTrashed ? i : -1).filter(i => i !== -1);
            } else {
                baseItemsPoolIndices = allImageData.map((item, i) => !item.isTrashed ? i : -1).filter(i => i !== -1);
            }

            displayedImageDataIndices = baseItemsPoolIndices.filter(index => {
                const item = allImageData[index];
                if (!item) return false; 

                let include = true;
                if (filterMode !== 'trashed') {
                    switch (filterMode) {
                        case 'tagged': if (item.tags.length === 0) include = false; break;
                        case 'untagged': if (item.tags.length > 0) include = false; break;
                        case 'modified': if (!item.modified) include = false; break;
                    }
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
        unifyPreviewOutput.innerHTML = ''; unifyApplyBtn.disabled = true; unifyStatus.textContent = '';
        convertPreviewOutput.innerHTML = ''; convertApplyBtn.disabled = true; convertStatus.textContent = '';
        currentSortProperty = 'path';
        sortPropertySelect.value = currentSortProperty;
        currentSortOrder = 'asc';
        sortOrderSelect.value = currentSortOrder;
        
        displayedImageDataIndices = getActiveItems().map(item => allImageData.indexOf(item));
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
        const item = allImageData[index];

        if (item.isTrashed) {
            editorStatus.textContent = "This item is in the trash. Restore it to edit tags.";
            editorStatus.style.color = 'var(--status-color)';
            editorTagsTextarea.disabled = true;
            editorSaveBtn.disabled = true;
            editorRemoveDupsBtn.disabled = true;
        } else {
            editorTagsTextarea.disabled = false;
            editorSaveBtn.disabled = false;
            editorRemoveDupsBtn.disabled = false;
            editorStatus.textContent = '';
        }

        currentEditIndex = index;
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
            const globalFocusedIndex = displayedImageDataIndices[keyboardFocusIndex];
            if (allImageData[globalFocusedIndex]) {
                 const focusedItemId = allImageData[globalFocusedIndex].id;
                 document.getElementById(focusedItemId)?.focus(); 
            }
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
        if (item.isTrashed) {
            editorStatus.textContent = "Cannot save tags for a trashed item.";
            editorStatus.style.color = 'var(--error-color)';
            return false;
        }

        const newTagsFromTextarea = parseTags(editorTagsTextarea.value);
        const newTagsStringFromTextarea = joinTags(newTagsFromTextarea);

        if (newTagsStringFromTextarea !== currentEditOriginalTagsString) {
            const oldTagsForStatUpdate = parseTags(currentEditOriginalTagsString); 
            
            item.tags = newTagsFromTextarea;
            item.modified = true;
            
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
            editorStatus.textContent = 'Tags updated. Remember to "Save All Changes".';
            editorStatus.style.color = 'var(--success-color)';
            currentEditOriginalTagsString = newTagsStringFromTextarea; 
            updateSaveAllButtonState();
            return true;
        } else {
            editorStatus.textContent = 'No changes detected to save.';
            editorStatus.style.color = 'var(--status-color)';
            return false;
        }
    }

    function removeDuplicateTagsInEditor() {
        if (currentEditIndex === -1 || !allImageData[currentEditIndex] || allImageData[currentEditIndex].isTrashed) return;
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
        if (item.isTrashed) return; 

        const oldLen = item.tags.length;
        const newTags = item.tags.filter(t => t !== tagToRemove);

        if (newTags.length < oldLen) {
            item.tags = newTags;
            item.modified = true;
            updateSaveAllButtonState();
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

    // --- Trash System Functions ---
    function promptMoveToTrash(itemId) {
        const item = allImageData[findIndexById(itemId)];
        if (!item) return;
        const msg = item.relativePath ? `Move "${item.relativePath}/${item.originalFullImageName}" to trash?` : `Move "${item.originalFullImageName}" to trash?`;
        if (confirm(`${msg}\n\nItem can be restored or permanently deleted from the trash.`)) {
            moveToTrash(itemId);
        }
    }

    function moveToTrash(itemId) {
        const index = findIndexById(itemId);
        if (index === -1) return;
        const item = allImageData[index];
        if (item.isTrashed) return; 

        item.isTrashed = true;
        item.modified = false; 

        const gridItemElement = document.getElementById(item.id);
        if (gridItemElement) {
            gridItemElement.classList.add('trashed-item');
            gridItemElement.classList.remove('modified'); 
            const actionBtn = gridItemElement.querySelector('.image-action-btn');
            if (actionBtn) {
                actionBtn.classList.remove('move-to-trash-btn');
                actionBtn.classList.add('restore-btn');
                actionBtn.textContent = '♻';
                actionBtn.title = `Restore "${item.originalFullImageName}" from trash`;
                actionBtn.replaceWith(actionBtn.cloneNode(true)); 
                gridItemElement.querySelector('.restore-btn').addEventListener('click', (e) => {
                    e.stopPropagation();
                    restoreFromTrash(item.id);
                });
            }
            const tagsDiv = gridItemElement.querySelector('.tags');
            if (tagsDiv) renderTagsForItem(tagsDiv, item); 
        }
        
        if (selectedItemIds.has(item.id)) {
            selectedItemIds.delete(item.id);
            gridItemElement?.classList.remove('selected');
            updateBatchEditButtonState();
        }
        if (lastInteractedItemId === item.id) lastInteractedItemId = null;

        recalculateAllTagStats(); 
        updateEmptyTrashButtonState();
        updateSaveAllButtonState(); 

        if (filterUntaggedSelect.value !== 'trashed') {
            filterAndSearch(); 
        } else {
            sortAndRender();
        }
        updateStatus(`Moved "${item.originalFullImageName}" to trash.`);
    }

    function restoreFromTrash(itemId) {
        const index = findIndexById(itemId);
        if (index === -1) return;
        const item = allImageData[index];
        if (!item.isTrashed) return; 

        item.isTrashed = false;

        const gridItemElement = document.getElementById(item.id);
        if (gridItemElement) {
            gridItemElement.classList.remove('trashed-item');
            if(item.modified) gridItemElement.classList.add('modified'); 

            const actionBtn = gridItemElement.querySelector('.image-action-btn');
            if (actionBtn) {
                actionBtn.classList.remove('restore-btn');
                actionBtn.classList.add('move-to-trash-btn');
                actionBtn.textContent = '×';
                const deleteTitle = item.relativePath ? `Move "${item.relativePath}/${item.originalFullImageName}" to trash` : `Move "${item.originalFullImageName}" to trash`;
                actionBtn.title = deleteTitle;
                actionBtn.replaceWith(actionBtn.cloneNode(true)); 
                gridItemElement.querySelector('.move-to-trash-btn').addEventListener('click', (e) => {
                    e.stopPropagation();
                    promptMoveToTrash(item.id);
                });
            }
             const tagsDiv = gridItemElement.querySelector('.tags');
            if (tagsDiv) renderTagsForItem(tagsDiv, item); 
        }

        recalculateAllTagStats(); 
        updateEmptyTrashButtonState();
        updateSaveAllButtonState();

        if (filterUntaggedSelect.value === 'trashed') {
            filterAndSearch();
        } else {
            sortAndRender();
        }
        updateStatus(`Restored "${item.originalFullImageName}" from trash.`);
    }

    async function permanentlyDeleteTrashedItems() {
        const itemsToDelete = allImageData.filter(item => item.isTrashed);
        if (itemsToDelete.length === 0) {
            updateStatus("Trash is already empty.");
            return;
        }
        if (!confirm(`Permanently delete ${itemsToDelete.length} item(s) from the trash? THIS CANNOT BE UNDONE.`)) {
            updateStatus("Permanent deletion cancelled.");
            return;
        }

        updateStatus(`Permanently deleting ${itemsToDelete.length} item(s)...`, true);
        let deletedCount = 0;
        let errorCount = 0;

        try {
            if (!datasetHandle || (await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted' &&
                await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted')) {
                updateStatus('Error: Write permission denied. Cannot delete files.');
                return;
            }

            for (const item of itemsToDelete) {
                let itemError = false;
                let parentDirHandle = datasetHandle;
                if (item.relativePath) { // Navigate to subdirectory if path exists
                    try {
                        const parts = item.relativePath.split('/');
                        for (const p of parts) {
                            if (p) parentDirHandle = await parentDirHandle.getDirectoryHandle(p, { create: false });
                        }
                    } catch (e) {
                        console.error(`Error accessing parent directory for ${item.originalFullImageName} in ${item.relativePath}:`, e);
                        itemError = true; // Cannot access parent, cannot delete
                    }
                }
                
                if (!itemError && item.imageHandle) {
                    try {
                        await parentDirHandle.removeEntry(item.originalFullImageName); 
                    } catch (e) {
                        console.error(`Error deleting image file ${item.originalFullImageName} from ${parentDirHandle.name}:`, e);
                        itemError = true;
                    }
                }
                if (!itemError && item.tagHandle) {
                     try {
                        await parentDirHandle.removeEntry(item.originalFullTagName); 
                    } catch (e) {
                        console.error(`Error deleting tag file ${item.originalFullTagName} from ${parentDirHandle.name}:`, e);
                        if (!itemError) itemError = true;
                    }
                }
                if (item.imageUrl) URL.revokeObjectURL(item.imageUrl);
                if (itemError) errorCount++; else deletedCount++;
            }
        } catch (permError) {
            updateStatus(`Permission error during deletion: ${permError.message}`);
        } finally {
            const newAllImageData = allImageData.filter(item => !item.isTrashed);
            allImageData = newAllImageData;
            
            clearSelection(); 
            clearKeyboardFocus();

            recalculateAllTagStats();
            updateEmptyTrashButtonState(); 
            updateSaveAllButtonState();

            filterAndSearch();

            updateStatus(`Permanently deleted ${deletedCount} item(s). ${errorCount > 0 ? `${errorCount} failed (see console).` : ''} Trash emptied.`, false);
        }
    }


    // --- Bulk Operations (Standard) ---
    function performStandardBulkOperation(operation) {
        const activeItems = getActiveItems();
        if (!datasetHandle || activeItems.length === 0) {
            updateStatus("Load dataset and ensure there are active (non-trashed) images.");
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
                    updateStatus("Invalid rename input."); return;
                }
                desc = `Rename "${oldTag}" to "${newTag}"`;
                break;
            case 'remove':
                toRemove = removeTagNameInput.value.trim();
                if (!toRemove) { updateStatus("Need tag to remove."); return; }
                desc = `Remove "${toRemove}"`; confirmNeeded = true;
                break;
            case 'add':
                toAdd = parseTags(addTriggerWordsInput.value);
                if (toAdd.length === 0) { updateStatus("Need tags to add."); return; }
                desc = `Add tags "${joinTags(toAdd)}"`;
                break;
            case 'removeDuplicates':
                desc = `Remove duplicate tags from each file`;
                break;
            default: updateStatus("Unknown standard bulk op."); return;
        }

        if (confirmNeeded && !confirm(`Bulk action on ALL ${activeItems.length} active images?\n\n${desc}`)) {
            updateStatus("Cancelled."); return;
        }

        updateStatus(`${desc} on active images...`, true);
        let statsChanged = false;
        activeItems.forEach((item) => { 
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
            if (itemChanged && operation !== 'removeDuplicates' && joinTags(currentTagsArray.slice().sort()) === joinTags(parseTags(origTagsString).slice().sort())) {
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
            updateSaveAllButtonState();
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
            updateStatus(`${modCount} active items affected by "${desc}". Save changes.`, false);
        } else {
            updateStatus(`Bulk op "${desc}" complete. No changes to active items.`, false);
        }
        if (operation === 'rename') { renameOldTagInput.value = ''; renameNewTagInput.value = ''; }
        if (operation === 'remove') removeTagNameInput.value = '';
        if (operation === 'add') addTriggerWordsInput.value = '';
    }

    // --- Actions from Tag Frequency List ---
    function handleDeleteTagFromFrequencyList(tagName) {
        if (!tagName) return;
        const activeItems = getActiveItems();
        if (activeItems.length === 0) {
            updateStatus("No active items to process."); return;
        }
        if (!confirm(`Are you sure you want to remove the tag "${tagName}" from ALL ${activeItems.length} active images where it appears?`)) {
            updateStatus("Tag deletion cancelled."); return;
        }

        updateStatus(`Removing tag "${tagName}" from active images...`, true);
        let modCount = 0;
        const affectedIds = new Set();

        activeItems.forEach(item => {
            const initialLength = item.tags.length;
            item.tags = item.tags.filter(t => t !== tagName);
            if (item.tags.length < initialLength) {
                item.modified = true;
                modCount++;
                affectedIds.add(item.id);
            }
        });

        if (modCount > 0) {
            updateSaveAllButtonState();
            affectedIds.forEach(id => {
                const gridItem = document.getElementById(id);
                const itemData = allImageData.find(i => i.id === id);
                if (gridItem && itemData) {
                    gridItem.classList.add('modified');
                    const tagsDiv = gridItem.querySelector('.tags');
                    if (tagsDiv) renderTagsForItem(tagsDiv, itemData);
                }
            });
            recalculateAllTagStats(); // This will also call updateStatisticsDisplay
            updateStatus(`${modCount} active items affected by removing tag "${tagName}". Remember to Save All Changes.`, false);
        } else {
            updateStatus(`Tag "${tagName}" was not found in any active items. No changes made.`, false);
            recalculateAllTagStats(); // Still good to refresh stats display in case something was off
        }
    }

    function handleRenameTagFromFrequencyList(oldTagName) {
        if (!oldTagName) return;
        const activeItems = getActiveItems();
        if (activeItems.length === 0) {
            updateStatus("No active items to process."); return;
        }
        const newTagName = prompt(`Rename all occurrences of "${oldTagName}" to:`, oldTagName);
        if (newTagName === null) { // User cancelled
            updateStatus("Tag rename cancelled."); return;
        }
        const trimmedNewTagName = newTagName.trim();
        if (!trimmedNewTagName || trimmedNewTagName === oldTagName) {
            updateStatus("Invalid new tag name or no change specified. Rename cancelled."); return;
        }
        if (!confirm(`Are you sure you want to rename the tag "${oldTagName}" to "${trimmedNewTagName}" in ALL ${activeItems.length} active images where it appears?`)) {
            updateStatus("Tag rename cancelled."); return;
        }

        updateStatus(`Renaming tag "${oldTagName}" to "${trimmedNewTagName}"...`, true);
        let modCount = 0;
        const affectedIds = new Set();

        activeItems.forEach(item => {
            if (item.tags.includes(oldTagName)) {
                item.tags = [...new Set(item.tags.map(t => t === oldTagName ? trimmedNewTagName : t))];
                item.modified = true;
                modCount++;
                affectedIds.add(item.id);
            }
        });

        if (modCount > 0) {
            updateSaveAllButtonState();
            affectedIds.forEach(id => {
                const gridItem = document.getElementById(id);
                const itemData = allImageData.find(i => i.id === id);
                if (gridItem && itemData) {
                    gridItem.classList.add('modified');
                    const tagsDiv = gridItem.querySelector('.tags');
                    if (tagsDiv) renderTagsForItem(tagsDiv, itemData);
                }
            });
            recalculateAllTagStats();
            updateStatus(`${modCount} active items affected by renaming tag. Remember to Save All Changes.`, false);
        } else {
            updateStatus(`Tag "${oldTagName}" was not found in any active items. No changes made.`, false);
            recalculateAllTagStats();
        }
    }


    // --- UTILITY FUNCTIONS (operate on active items by default) ---
    function findDuplicateTagFiles() {
        const activeItems = getActiveItems();
        if (activeItems.length < 2) {
            duplicateFilesOutput.innerHTML = '<p>Need at least 2 active files to compare.</p>';
            return;
        }
        updateStatus('Finding files with identical tags (among active files)...', true);
        const tagStringToFilesMap = new Map();
        activeItems.forEach(item => {
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
                duplicateSets.push({ tags: tagStr, files: files.sort() });
            }
        }
        duplicateSets.sort((a, b) => a.tags.localeCompare(b.tags));
        duplicateFilesOutput.innerHTML = '';
        if (duplicateSets.length > 0) {
            const summaryP = document.createElement('p');
            summaryP.textContent = `Found ${duplicateSets.length} set(s) of active files with identical tags:`;
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
            updateStatus(`Found ${duplicateSets.length} duplicate set(s) among active files.`, false);
        } else {
            duplicateFilesOutput.innerHTML = '<p>No active files found with identical tags.</p>';
            updateStatus('No duplicate tag sets found among active files.', false);
        }
    }

    async function findOrphanedAndMissingTagFiles() { 
        if (!datasetHandle) {
            orphanedMissingOutput.innerHTML = '<p>Load dataset first.</p>'; return;
        }
        updateStatus('Checking for orphaned/missing tag files...', true);
        const orphanedTagFiles = []; 
        const imagesMissingTagFiles = []; 
        const allLoadedImageBaseKeys = new Set(); // Key: relativePath/imageName (base name without ext)
        const allLoadedTagFileFullNames = new Set(); // Key: relativePath/originalFullTagName

        allImageData.forEach(item => {
            const imgBaseKey = item.relativePath ? `${item.relativePath}/${item.imageName}` : item.imageName;
            allLoadedImageBaseKeys.add(imgBaseKey);
            if (item.tagHandle) {
                const tagFullNameKey = item.relativePath ? `${item.relativePath}/${item.originalFullTagName}` : item.originalFullTagName;
                allLoadedTagFileFullNames.add(tagFullNameKey);
            } else if (!item.isTrashed) { 
                 imagesMissingTagFiles.push(item.relativePath ? `${item.relativePath}/${item.originalFullImageName}` : item.originalFullImageName);
            }
        });
        
        for (const [fullPathOnDisk, entryInfo] of allFileSystemEntries.entries()) {
            if (entryInfo.kind !== 'file' || !entryInfo.name.endsWith(TAG_EXTENSION)) continue;

            const diskFileKey = fullPathOnDisk.substring(0, fullPathOnDisk.lastIndexOf(TAG_EXTENSION)); // e.g. "sub/image.png" or "sub/image"
            const diskFileKeyWithoutImgExt = diskFileKey.substring(0, diskFileKey.lastIndexOf('.')) !== -1 ? diskFileKey.substring(0, diskFileKey.lastIndexOf('.')) : diskFileKey;

            // An orphaned tag file could be:
            // 1. "image.txt" where no "image.any_image_ext" exists.
            // 2. "image.png.txt" where no "image.png" exists.
            if (!allLoadedImageBaseKeys.has(diskFileKey) && !allLoadedImageBaseKeys.has(diskFileKeyWithoutImgExt)) { 
                orphanedTagFiles.push(fullPathOnDisk);
            }
        }

        orphanedMissingOutput.innerHTML = '';
        let foundAny = false;
        if (imagesMissingTagFiles.length > 0) {
            foundAny = true;
            const p = document.createElement('p');
            p.textContent = `Active images missing expected tag files (e.g., basename.txt) (${imagesMissingTagFiles.length}):`;
            orphanedMissingOutput.appendChild(p);
            const ul = document.createElement('ul');
            imagesMissingTagFiles.sort().forEach(filePath => {
                const li = document.createElement('li'); li.textContent = escapeHtml(filePath); ul.appendChild(li);
            });
            orphanedMissingOutput.appendChild(ul);
        }
        if (orphanedTagFiles.length > 0) {
            foundAny = true;
            const p = document.createElement('p');
            p.textContent = `Orphaned tag files on disk (no matching image loaded or tag filename mismatch) (${orphanedTagFiles.length}):`;
            orphanedMissingOutput.appendChild(p);
            const ul = document.createElement('ul');
            orphanedTagFiles.sort().forEach(filePath => {
                const li = document.createElement('li'); li.textContent = escapeHtml(filePath); ul.appendChild(li);
            });
            orphanedMissingOutput.appendChild(ul);
        }
        if (!foundAny) {
            orphanedMissingOutput.innerHTML = '<p>No orphaned or missing tag files found according to current data load and naming conventions.</p>';
        }
        updateStatus('Orphaned/missing tag file check complete.', false);
    }

    function findCaseInconsistentTags() { 
        if (uniqueTags.size < 2) {
            caseInconsistentOutput.innerHTML = '<p>Not enough unique tags from active items to compare.</p>';
            return;
        }
        updateStatus('Finding case-inconsistent tags (among active items)...', true);
        const tagMap = new Map(); 
        uniqueTags.forEach(tag => {
            const lower = tag.toLowerCase();
            if (!tagMap.has(lower)) tagMap.set(lower, []);
            tagMap.get(lower).push(tag);
        });
        const inconsistentGroups = [];
        for (const [lower, originals] of tagMap.entries()) {
            if (originals.length > 1) {
                const uniqueOriginals = new Set(originals);
                if (uniqueOriginals.size > 1) { 
                    inconsistentGroups.push({ lower, originals: [...uniqueOriginals].sort() });
                }
            }
        }
        inconsistentGroups.sort((a,b) => a.lower.localeCompare(b.lower));
        caseInconsistentOutput.innerHTML = '';
        if (inconsistentGroups.length > 0) {
            const summaryP = document.createElement('p');
            summaryP.textContent = `Found ${inconsistentGroups.length} group(s) of case-inconsistent tags (from active items):`;
            caseInconsistentOutput.appendChild(summaryP);
            inconsistentGroups.forEach(group => {
                const div = document.createElement('div'); div.classList.add('tag-group');
                const strong = document.createElement('strong');
                strong.textContent = `Base: "${escapeHtml(group.lower)}" - Variants:`; div.appendChild(strong);
                const ul = document.createElement('ul');
                group.originals.forEach(originalTag => {
                    const li = document.createElement('li'); li.textContent = escapeHtml(originalTag); ul.appendChild(li);
                });
                div.appendChild(ul); caseInconsistentOutput.appendChild(div);
            });
            updateStatus(`Found ${inconsistentGroups.length} case-inconsistent tag group(s) from active items.`, false);
        } else {
            caseInconsistentOutput.innerHTML = '<p>No case-inconsistent tags found among active items.</p>';
            updateStatus('No case-inconsistent tags found among active items.', false);
        }
    }

    function exportUniqueTagList() { 
        if (uniqueTags.size === 0) {
            updateStatus("No unique tags from active items to export."); return;
        }
        const tagArray = [...uniqueTags].sort();
        const fileContent = tagArray.join('\n');
        const blob = new Blob([fileContent], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'unique_tags_export.txt';
        document.body.appendChild(a); a.click();
        document.body.removeChild(a); URL.revokeObjectURL(url);
        updateStatus(`Exported ${tagArray.length} unique tags from active items.`);
    }


    async function saveAllChanges() { 
        if (!datasetHandle) {
            updateStatus("No dataset loaded."); return;
        }
        const itemsToSave = getActiveItems().filter(item => item.modified);
        if (itemsToSave.length === 0) {
            updateStatus("No tag modifications on active items to save.");
            updateSaveAllButtonState(); 
            return;
        }
        let savedCount = 0, errorCount = 0;
        updateStatus(`Saving ${itemsToSave.length} modified tag file(s) for active items...`, true);
        try {
            if (await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted' &&
                await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted') {
                updateStatus('Error: Write permission denied. Cannot save tags.');
                return; 
            }
        } catch (permError) {
            updateStatus(`Error requesting permission: ${permError.message}`);
            return;
        }

        for (const item of itemsToSave) { 
            const newTagString = joinTags(item.tags);
            // Determine the correct tag filename. If originalFullTagName exists and is specific (like image.png.txt), use that pattern.
            // Otherwise, default to imageName.txt.
            let tagFileNameToUse;
            if (item.originalFullTagName && item.originalFullTagName.startsWith(item.imageName + item.originalFullImageName.substring(item.originalFullImageName.lastIndexOf('.')))) {
                // Kohya-style: e.g. if image is "basename.png", tag is "basename.png.txt"
                 tagFileNameToUse = item.imageName + item.originalFullImageName.substring(item.originalFullImageName.lastIndexOf('.')) + TAG_EXTENSION;
            } else {
                // Default: "basename.txt"
                tagFileNameToUse = item.imageName + TAG_EXTENSION;
            }


            const fullLogPath = item.relativePath ? `${item.relativePath}/${tagFileNameToUse}` : tagFileNameToUse;
            try {
                let tagHandle = item.tagHandle;
                if (!tagHandle || (item.originalFullTagName && item.originalFullTagName !== tagFileNameToUse) ) { 
                    // Need to get or create the handle, especially if the expected name changed
                    try {
                        let parentDirHandle = datasetHandle;
                        if (item.relativePath) {
                            const pathParts = item.relativePath.split('/');
                            for (const part of pathParts) {
                                if (part) parentDirHandle = await parentDirHandle.getDirectoryHandle(part, { create: false });
                            }
                        }
                        // If an old handle exists but name is different, remove old before creating new to avoid issues.
                        // This is tricky if it was renamed by unification. Best if saveAllChanges *always* uses the current expected name.
                        tagHandle = await parentDirHandle.getFileHandle(tagFileNameToUse, { create: true });
                        item.tagHandle = tagHandle;
                        item.originalFullTagName = tagFileNameToUse; // Update our record of the tag name
                    } catch (handleError) {
                        console.error(`Error getting/creating tag file handle for ${fullLogPath}:`, handleError);
                        errorCount++; item.modified = true; 
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
                errorCount++; item.modified = true; 
            }
        }
        let finalMessage = `Saved ${savedCount} tag file(s).`;
        if (errorCount > 0) finalMessage += ` Failed to save ${errorCount}. Check console.`;
        
        updateSaveAllButtonState(); 
        recalculateAllTagStats(); 
        updateStatus(finalMessage, false);
    }

    function navigateEditor(direction) { 
        if (currentEditIndex === -1 || displayedImageDataIndices.length <= 1) return;

        const currentGlobalIndex = currentEditIndex; 
        let currentDisplayIndexOfEditedItem = -1;
        for(let i=0; i < displayedImageDataIndices.length; i++){
            if(displayedImageDataIndices[i] === currentGlobalIndex){
                currentDisplayIndexOfEditedItem = i;
                break;
            }
        }
        
        if(currentDisplayIndexOfEditedItem === -1) {
             console.warn("Edited item not in current display for navigation. Closing editor.");
             closeEditor();
             return;
        }

        let nextItemDisplayIndex;
        if (direction === 'next') {
            nextItemDisplayIndex = (currentDisplayIndexOfEditedItem + 1) % displayedImageDataIndices.length;
        } else { 
            nextItemDisplayIndex = (currentDisplayIndexOfEditedItem - 1 + displayedImageDataIndices.length) % displayedImageDataIndices.length;
        }
        
        const itemBeingEdited = allImageData[currentGlobalIndex];
        if (!itemBeingEdited.isTrashed && editorTagsTextarea.value !== currentEditOriginalTagsString) {
            saveEditorChanges(); 
        }

        const nextGlobalIndexToOpen = displayedImageDataIndices[nextItemDisplayIndex];
        if (allImageData[nextGlobalIndexToOpen]) {
            openEditor(allImageData[nextGlobalIndexToOpen].id);
        }
    }

    function updateEditorNavButtonStates() {
        const currentItem = (currentEditIndex !== -1) ? allImageData[currentEditIndex] : null;
        const disableNav = displayedImageDataIndices.length <= 1 || (currentItem && currentItem.isTrashed);

        editorPrevBtn.disabled = disableNav;
        editorNextBtn.disabled = disableNav;
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
        
        currentSuggestions = [...uniqueTags] 
            .filter(tag => 
                tag.toLowerCase().includes(currentTagFragment.toLowerCase()) && 
                !existingTagsInInput.includes(tag)
            )
            .sort((a,b) => a.localeCompare(b)) 
            .slice(0, 10); 

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
            items[activeSuggestionIndex].scrollIntoView({ block: 'nearest' });
        }
    }

    function openBatchEditor() {
        if (selectedItemIds.size === 0) {
            updateStatus("No items selected for batch editing.");
            return;
        }
        const activeSelectedItemsData = [];
        selectedItemIds.forEach(id => {
            const item = getActiveItemById(id); 
            if (item) activeSelectedItemsData.push(item);
        });

        if (activeSelectedItemsData.length !== selectedItemIds.size) {
            updateStatus("Some selected items are trashed and cannot be batch edited. Deselect them or restore them.", false);
            return;
        }
        if (activeSelectedItemsData.length === 0) {
            updateStatus("No active items selected for batch editing.");
            return;
        }


        batchEditorModal.classList.remove('hidden');
        batchEditorSelectedCount.textContent = activeSelectedItemsData.length;
        batchEditorAddTagsInput.value = '';
        batchEditorRemoveTagsInput.value = '';
        batchEditorStatus.textContent = '';

        let commonTags = new Set(activeSelectedItemsData[0].tags);
        const allTagsInSelection = new Map(); 

        activeSelectedItemsData.forEach((item, idx) => {
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
                const span = document.createElement('span'); span.classList.add('tag'); span.textContent = tag;
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
                const span = document.createElement('span'); span.classList.add('tag'); span.textContent = tag;
                const countSpan = document.createElement('span'); countSpan.classList.add('count');
                countSpan.textContent = `(${count}/${activeSelectedItemsData.length})`; span.appendChild(countSpan);
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
            if (item.isTrashed) return; 

            const originalItemTagsString = joinTags(item.tags);
            let currentItemTagsSet = new Set(item.tags);

            tagsToAdd.forEach(tag => currentItemTagsSet.add(tag));
            tagsToRemove.forEach(tag => currentItemTagsSet.delete(tag));
            
            const newItemTagsArray = [...currentItemTagsSet].sort(); 
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
            updateSaveAllButtonState();
            if (overallStatsChanged) recalculateAllTagStats(); 
            batchEditorStatus.textContent = `Tag changes applied to ${modifiedCount} item(s). Remember to "Save All Changes".`;
            batchEditorStatus.style.color = 'var(--success-color)';
            batchEditorAddTagsInput.value = ''; 
            batchEditorRemoveTagsInput.value = '';
        } else {
            batchEditorStatus.textContent = "No changes made to selected items (tags might already be present/absent).";
            batchEditorStatus.style.color = 'var(--status-color)';
        }
    }
    
    function batchMoveSelectedToTrash() {
        if (selectedItemIds.size === 0) {
            batchEditorStatus.textContent = "No items selected to move to trash.";
            batchEditorStatus.style.color = 'var(--error-color)';
            return;
        }
        if (confirm(`Move ${selectedItemIds.size} selected item(s) to trash?`)) {
            let movedCount = 0;
            const idsToTrash = new Set(selectedItemIds); 
            idsToTrash.forEach(itemId => {
                const item = allImageData.find(i => i.id === itemId);
                if (item && !item.isTrashed) {
                    moveToTrash(itemId); 
                    movedCount++;
                }
            });

            if (movedCount > 0) {
                updateStatus(`Moved ${movedCount} item(s) to trash.`);
            } else {
                updateStatus("No active items were moved to trash (they might have been already trashed).");
            }
            closeBatchEditor();
            filterAndSearch(); 
        } else {
            batchEditorStatus.textContent = "Move to trash cancelled.";
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
                id: generateRuleId('default-replace-underscores'), name: "Default: Replace Underscores with Spaces",
                find: "_", replace: " ", isRegex: false, isCaseSensitive: false, applyTo: "all", specificTags: []
            });
            customRules.push({
                id: generateRuleId('default-escape-parentheses'), name: "Default: Escape Parentheses () -> \\(\\) ",
                find: "(\\(|\\))", replace: "\\\\$1", isRegex: true, isCaseSensitive: true, applyTo: "all", specificTags: []
            });
            customRules.push({
                id: generateRuleId('default-smart-brackets'), name: "Default: Format 'word_(info)' to 'word from info'",
                find: "^([^\\s_()]+)(?:_)?\\(([^)]+)\\)$", replace: "$1 from $2", isRegex: true, isCaseSensitive: false, applyTo: "all", specificTags: []
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
            customRulesListDiv.innerHTML = '<p>No custom rules defined yet.</p>'; return;
        }
        const ul = document.createElement('ul'); ul.classList.add('custom-rule-item-list');
        customRules.forEach(rule => {
            const li = document.createElement('li'); li.dataset.ruleId = rule.id;
            const nameSpan = document.createElement('span'); nameSpan.classList.add('rule-name'); nameSpan.textContent = rule.name; li.appendChild(nameSpan);
            const summarySpan = document.createElement('span'); summarySpan.classList.add('rule-summary');
            let findText = rule.find.length > 20 ? rule.find.substring(0, 17) + "..." : rule.find;
            let replaceText = rule.replace.length > 20 ? rule.replace.substring(0, 17) + "..." : rule.replace;
            summarySpan.textContent = `Find: "${escapeHtml(findText)}", Replace: "${escapeHtml(replaceText)}" ${rule.isRegex ? '(Regex)' : ''}`; li.appendChild(summarySpan);
            const actionsDiv = document.createElement('div'); actionsDiv.classList.add('rule-actions');
            const editBtn = document.createElement('button'); editBtn.textContent = 'Edit'; editBtn.classList.add('edit-rule-btn');
            editBtn.addEventListener('click', () => openRuleEditor(rule.id)); actionsDiv.appendChild(editBtn);
            const deleteBtn = document.createElement('button'); deleteBtn.textContent = 'Delete'; deleteBtn.classList.add('delete-rule-btn');
            deleteBtn.addEventListener('click', () => deleteCustomRule(rule.id)); actionsDiv.appendChild(deleteBtn);
            li.appendChild(actionsDiv); ul.appendChild(li);
        });
        customRulesListDiv.appendChild(ul);
    }
    function populateCustomRulesBulkApplyDropdown() {
        bulkApplyCustomRulesSelect.innerHTML = '';
        if (customRules.length === 0) {
            const option = document.createElement('option'); option.value = ""; option.textContent = "No custom rules defined"; option.disabled = true;
            bulkApplyCustomRulesSelect.appendChild(option);
            bulkApplyCustomRulesBtn.disabled = true;
        } else {
            customRules.forEach(rule => {
                const option = document.createElement('option'); option.value = rule.id; option.textContent = rule.name;
                bulkApplyCustomRulesSelect.appendChild(option);
            });
            bulkApplyCustomRulesBtn.disabled = false;
        }
    }
    function openRuleEditor(ruleId = null) {
        editingRuleId = ruleId;
        ruleEditorModal.classList.remove('hidden');
        ruleEditorStatus.textContent = ''; ruleSpecificTagsGroup.style.display = 'none';
        if (ruleId) {
            const rule = customRules.find(r => r.id === ruleId);
            if (!rule) { ruleEditorStatus.textContent = 'Error: Rule not found.'; ruleEditorStatus.style.color = 'var(--error-color)'; return; }
            ruleEditorTitle.textContent = 'Edit Custom Rule'; ruleIdInput.value = rule.id;
            ruleNameInput.value = rule.name; ruleFindInput.value = rule.find; ruleReplaceInput.value = rule.replace;
            ruleIsRegexCheckbox.checked = rule.isRegex;
            ruleIsCaseSensitiveCheckbox.checked = rule.isRegex ? rule.isCaseSensitive : false;
            ruleIsCaseSensitiveCheckbox.disabled = !rule.isRegex;
            ruleApplyToSelect.value = rule.applyTo;
            if (rule.applyTo === 'only_if_is' || rule.applyTo === 'not_if_is') {
                ruleSpecificTagsGroup.style.display = 'block';
                ruleSpecificTagsInput.value = rule.specificTags ? rule.specificTags.join(', ') : '';
            }
        } else {
            ruleEditorTitle.textContent = 'Create New Custom Rule'; ruleIdInput.value = generateRuleId();
            ruleNameInput.value = ''; ruleFindInput.value = ''; ruleReplaceInput.value = '';
            ruleIsRegexCheckbox.checked = false; ruleIsCaseSensitiveCheckbox.checked = false; ruleIsCaseSensitiveCheckbox.disabled = true;
            ruleApplyToSelect.value = 'all'; ruleSpecificTagsInput.value = '';
        }
        ruleNameInput.focus();
    }
    function closeRuleEditor() { ruleEditorModal.classList.add('hidden'); editingRuleId = null; }
    ruleIsRegexCheckbox.addEventListener('change', () => {
        ruleIsCaseSensitiveCheckbox.disabled = !ruleIsRegexCheckbox.checked;
        if (!ruleIsRegexCheckbox.checked) ruleIsCaseSensitiveCheckbox.checked = false;
    });
    ruleApplyToSelect.addEventListener('change', () => {
        ruleSpecificTagsGroup.style.display = (ruleApplyToSelect.value === 'only_if_is' || ruleApplyToSelect.value === 'not_if_is') ? 'block' : 'none';
    });
    ruleEditorSaveBtn.addEventListener('click', () => {
        const id = ruleIdInput.value; const name = ruleNameInput.value.trim();
        const findPattern = ruleFindInput.value; const replacePattern = ruleReplaceInput.value;
        const isRegex = ruleIsRegexCheckbox.checked; const isCaseSensitive = ruleIsCaseSensitiveCheckbox.checked;
        const applyTo = ruleApplyToSelect.value; const specificTags = parseTags(ruleSpecificTagsInput.value);
        if (!name) { ruleEditorStatus.textContent = 'Rule name is required.'; ruleEditorStatus.style.color = 'var(--error-color)'; return; }
        if (!findPattern && !isRegex) { ruleEditorStatus.textContent = 'Find pattern required for non-regex.'; ruleEditorStatus.style.color = 'var(--error-color)'; return; }
        if ((applyTo === 'only_if_is' || applyTo === 'not_if_is') && specificTags.length === 0) {
            ruleEditorStatus.textContent = 'Specific tags required for this "Apply To" condition.'; ruleEditorStatus.style.color = 'var(--error-color)'; return;
        }
        if (isRegex) { try { new RegExp(findPattern); } catch (e) { ruleEditorStatus.textContent = `Invalid Regex: ${e.message}`; ruleEditorStatus.style.color = 'var(--error-color)'; return; } }
        const ruleData = { id, name, find: findPattern, replace: replacePattern, isRegex, isCaseSensitive, applyTo, specificTags };
        const index = editingRuleId ? customRules.findIndex(r => r.id === editingRuleId) : -1;
        if (index !== -1) customRules[index] = ruleData; else customRules.push(ruleData);
        saveCustomRules(); renderCustomRulesList(); populateCustomRulesBulkApplyDropdown(); closeRuleEditor();
        updateStatus(`Custom rule "${name}" ${editingRuleId ? 'updated' : 'created'}.`);
    });
    function deleteCustomRule(ruleId) {
        if (confirm(`Delete rule: "${customRules.find(r=>r.id===ruleId)?.name || 'this rule'}"?`)) {
            customRules = customRules.filter(r => r.id !== ruleId);
            saveCustomRules(); renderCustomRulesList(); populateCustomRulesBulkApplyDropdown();
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
                } catch (e) { console.warn(`Regex error rule "${rule.name}" tag "${tag}": ${e.message}`); return tag; }
            } else {
                if (rule.find === '') return newTag; 
                if (rule.isCaseSensitive) { newTag = newTag.split(rule.find).join(rule.replace); } 
                else { const escFind = rule.find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); const regex = new RegExp(escFind, 'gi'); newTag = newTag.replace(regex, rule.replace); }
            }
            return newTag;
        };
        switch (rule.applyTo) {
            case 'all': return applyThisRule();
            case 'only_if_is': if (rule.specificTags.includes(tag)) return applyThisRule(); break;
            case 'not_if_is': if (!rule.specificTags.includes(tag)) return applyThisRule(); break;
        }
        return tag;
    }
    function performBulkCustomRulesOperation() { 
        const selectedRuleIds = Array.from(bulkApplyCustomRulesSelect.selectedOptions).map(opt => opt.value);
        if (selectedRuleIds.length === 0) { updateStatus("No custom rules selected."); return; }
        const rulesToApply = selectedRuleIds.map(id => customRules.find(r => r.id === id)).filter(r => r);
        if (rulesToApply.length === 0) { updateStatus("Selected rules not found."); return; }
        
        const activeItems = getActiveItems();
        if (activeItems.length === 0) { updateStatus("No active images to apply rules to."); return; }

        if (!confirm(`Apply ${rulesToApply.length} rule(s) to ALL ${activeItems.length} active images?`)) {
            updateStatus("Rule application cancelled."); return;
        }
        updateStatus(`Applying ${rulesToApply.length} custom rule(s) to active images...`, true);
        let modCount = 0; let affectedItemIds = new Set(); let overallStatsChanged = false;

        activeItems.forEach(item => { 
            let originalItemTagsString = joinTags(item.tags);
            let currentItemTagsArray = [...item.tags];
            rulesToApply.forEach(rule => {
                currentItemTagsArray = currentItemTagsArray.map(tag => applySingleCustomRuleToTag(tag, rule));
            });
            currentItemTagsArray = [...new Set(currentItemTagsArray.map(t => t.trim()).filter(t => t))]; 
            if (joinTags(currentItemTagsArray) !== originalItemTagsString) {
                item.tags = currentItemTagsArray; item.modified = true;
                modCount++; affectedItemIds.add(item.id); overallStatsChanged = true;
            }
        });
        if (modCount > 0) {
            updateSaveAllButtonState();
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
            updateStatus(`${modCount} active items affected. Save changes.`, false);
        } else {
            updateStatus(`Rules applied. No changes to active items.`, false);
        }
    }

    // --- Dataset Unification Logic ---
    function toggleUnifyPaddingLengthGroup() {
        unifyPaddingLengthGroup.style.display = unifyPadZerosCheckbox.checked ? 'block' : 'none';
    }

    function generateUnificationPreview() {
        unifyPreviewOutput.innerHTML = '';
        unifyApplyBtn.disabled = true;
        unifyStatus.textContent = '';
        unificationPreviewData = [];

        const activeItems = getActiveItems();
        if (activeItems.length === 0) {
            unifyStatus.textContent = 'No active images to unify.';
            unifyStatus.style.color = 'var(--status-color)';
            return;
        }

        const prefix = unifyPrefixInput.value.trim();
        let startNum = parseInt(unifyStartNumInput.value);
        const padZeros = unifyPadZerosCheckbox.checked;
        const paddingLength = parseInt(unifyPaddingLengthInput.value);

        if (isNaN(startNum) || startNum < 0) {
            unifyStatus.textContent = 'Invalid start number.';
            unifyStatus.style.color = 'var(--error-color)';
            return;
        }
        if (padZeros && (isNaN(paddingLength) || paddingLength < 1 || paddingLength > 10)) {
            unifyStatus.textContent = 'Invalid padding length (must be 1-10).';
            unifyStatus.style.color = 'var(--error-color)';
            return;
        }

        const outputFragment = document.createDocumentFragment();
        const newFilenamesSet = new Set();
        let hasChanges = false;
        let currentNum = startNum;

        activeItems.forEach(item => {
            const numberStr = padZeros ? String(currentNum).padStart(paddingLength, '0') : String(currentNum);
            const originalExt = item.originalFullImageName.substring(item.originalFullImageName.lastIndexOf('.'));
            const newBaseName = `${prefix}${numberStr}`;
            const newImageName = `${newBaseName}${originalExt}`; // New image name maintains original extension
            
            let newTagName = null;
            if (item.originalFullTagName) { // If a tag file existed
                 // If original tag was like "base.ext.txt", new one should be "newBase.ext.txt"
                if (item.originalFullTagName === item.imageName + originalExt + TAG_EXTENSION) {
                    newTagName = newBaseName + originalExt + TAG_EXTENSION;
                } 
                // If original tag was like "base.txt", new one should be "newBase.txt"
                else if (item.originalFullTagName === item.imageName + TAG_EXTENSION) {
                    newTagName = newBaseName + TAG_EXTENSION;
                } else {
                    // Unknown or complex original tag naming, try to adapt based on new base name
                    newTagName = newBaseName + TAG_EXTENSION; // Default to newBase.txt
                    console.warn(`Unusual original tag name "${item.originalFullTagName}" for image "${item.originalFullImageName}". Defaulting new tag to "${newTagName}".`);
                }
            }


            const oldDisplayPath = item.relativePath ? `${item.relativePath}/${item.originalFullImageName}` : item.originalFullImageName;
            const newDisplayPath = item.relativePath ? `${item.relativePath}/${newImageName}` : newImageName;

            let changeEntry = `<p><strong>${escapeHtml(oldDisplayPath)}</strong> → <strong>${escapeHtml(newDisplayPath)}</strong>`;
            if (item.originalFullTagName && newTagName) {
                 const oldTagDisplayPath = item.relativePath ? `${item.relativePath}/${item.originalFullTagName}` : item.originalFullTagName;
                 const newTagDisplayPath = item.relativePath ? `${item.relativePath}/${newTagName}` : newTagName;
                 if (item.originalFullTagName !== newTagName) { // Only show if tag name actually changes
                    changeEntry += `<br>&nbsp;&nbsp;&nbsp;<em>${escapeHtml(oldTagDisplayPath)}</em> → <em>${escapeHtml(newTagDisplayPath)}</em>`;
                 }
            }
            changeEntry += `</p>`;
            const p = document.createElement('div'); 
            p.innerHTML = changeEntry;
            outputFragment.appendChild(p);

            if (item.originalFullImageName !== newImageName || (item.originalFullTagName && newTagName && item.originalFullTagName !== newTagName)) {
                hasChanges = true;
            }

            unificationPreviewData.push({
                itemId: item.id,
                oldImageName: item.originalFullImageName,
                newImageName: newImageName,
                oldTagName: item.originalFullTagName, 
                newTagName: newTagName,             
                relativePath: item.relativePath,
            });

            const fullNewPathKey = item.relativePath ? `${item.relativePath}/${newImageName}` : newImageName;
            if (newFilenamesSet.has(fullNewPathKey)) {
                 unifyStatus.textContent = `Error: Generated filename collision for "${escapeHtml(fullNewPathKey)}". Adjust prefix, start number, or padding.`;
                 unifyStatus.style.color = 'var(--error-color)';
                 unifyPreviewOutput.innerHTML = ''; 
                 unificationPreviewData = []; 
                 return; 
            }
            newFilenamesSet.add(fullNewPathKey);
            currentNum++;
        });

        if (unifyStatus.textContent.startsWith('Error:')) return; 

        if (unificationPreviewData.length > 0) {
            unifyPreviewOutput.appendChild(outputFragment);
            if (hasChanges) {
                unifyApplyBtn.disabled = false;
                unifyStatus.textContent = `Previewing ${unificationPreviewData.length} potential renames. Please verify.`;
                unifyStatus.style.color = 'var(--status-color)';
            } else {
                unifyStatus.textContent = 'No filename changes based on current settings.';
                unifyStatus.style.color = 'var(--status-color)';
            }
        } else {
            unifyStatus.textContent = 'No active items to preview for unification.';
            unifyStatus.style.color = 'var(--status-color)';
        }
    }

    async function applyUnification() {
        if (unificationPreviewData.length === 0 || unifyApplyBtn.disabled) {
            unifyStatus.textContent = 'No changes to apply or preview first.';
            unifyStatus.style.color = 'var(--error-color)';
            return;
        }

        if (!confirm(`This will rename ${unificationPreviewData.length} file(s) on your disk based on the preview. This action is IRREVERSIBLE. Are you sure you want to proceed?`)) {
            unifyStatus.textContent = 'Unification cancelled by user.';
            unifyStatus.style.color = 'var(--status-color)';
            return;
        }

        updateStatus('Applying unification renames...', true);
        unifyApplyBtn.disabled = true;
        let successCount = 0;
        let errorCount = 0;

        try {
             if (!datasetHandle || (await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted' &&
                await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted')) {
                updateStatus('Error: Write permission denied. Cannot rename files.');
                unifyStatus.textContent = 'Write permission denied.';
                unifyStatus.style.color = 'var(--error-color)';
                return;
            }

            for (const change of unificationPreviewData) {
                const item = allImageData.find(i => i.id === change.itemId);
                if (!item) { errorCount++; console.warn(`Item ${change.itemId} not found during apply unification.`); continue; }
                if (item.isTrashed) { console.warn(`Skipping trashed item ${change.itemId} for unification.`); continue; }

                let parentDirHandle = datasetHandle;
                if (item.relativePath) {
                    try {
                        const parts = item.relativePath.split('/');
                        for (const p of parts) {
                            if (p) parentDirHandle = await parentDirHandle.getDirectoryHandle(p, { create: false });
                        }
                    } catch (e) {
                        console.error(`Error navigating to directory ${item.relativePath} for ${item.originalFullImageName}: ${e}`);
                        errorCount++; continue;
                    }
                }
                
                let itemRenameError = false;
                // 1. Rename Image File
                if (item.imageHandle && change.oldImageName !== change.newImageName) {
                    try {
                        await item.imageHandle.move(parentDirHandle, change.newImageName); 
                                                
                        item.imageHandle = await parentDirHandle.getFileHandle(change.newImageName); 
                        if (item.imageUrl) URL.revokeObjectURL(item.imageUrl);
                        const newFile = await item.imageHandle.getFile();
                        item.imageUrl = URL.createObjectURL(newFile);
                        item.originalFullImageName = change.newImageName;
                        item.imageName = change.newImageName.substring(0, change.newImageName.lastIndexOf('.'));
                    } catch (e) {
                        console.error(`Error renaming image "${change.oldImageName}" to "${change.newImageName}" in ${parentDirHandle.name}: ${e}`);
                        errorCount++; itemRenameError = true;
                    }
                }

                // 2. Rename Tag File
                if (!itemRenameError && item.tagHandle && change.oldTagName && change.newTagName && change.oldTagName !== change.newTagName) {
                    try {
                        await item.tagHandle.move(parentDirHandle, change.newTagName); 

                        item.tagHandle = await parentDirHandle.getFileHandle(change.newTagName); 
                        item.originalFullTagName = change.newTagName;
                    } catch (e) {
                        console.error(`Error renaming tag "${change.oldTagName}" to "${change.newTagName}" in ${parentDirHandle.name}: ${e}`);
                        errorCount++; 
                    }
                } else if (!itemRenameError && !item.tagHandle && change.newTagName && item.originalFullImageName !== change.newImageName) {
                    // If image was renamed and there was NO original tag file, we still update the expected tag name.
                    // This is mostly for data consistency if a tag file *were* to be created later for the new image name.
                    item.originalFullTagName = change.newTagName; // This might set it to null if change.newTagName is null
                }


                if (!itemRenameError) { 
                    item.modified = false; 
                    document.getElementById(item.id)?.classList.remove('modified');
                    successCount++;
                }
            }

        } catch (e) {
            updateStatus(`Error during unification: ${e.message}`, false);
            unifyStatus.textContent = `Error: ${e.message}`;
            unifyStatus.style.color = 'var(--error-color)';
        } finally {
            unifyPreviewOutput.innerHTML = '';
            unificationPreviewData = [];
            unifyApplyBtn.disabled = true;
            
            let finalMsg = `Unification complete. ${successCount} item(s) processed successfully.`;
            if (errorCount > 0) finalMsg += ` ${errorCount} errors occurred (see console).`;
            unifyStatus.textContent = finalMsg;
            unifyStatus.style.color = errorCount > 0 ? 'var(--error-color)' : 'var(--success-color)';
            
            updateStatus(finalMsg, false);
            filterAndSearch(); 
            recalculateAllTagStats(); 
        }
    }

    // --- Image Format Conversion ---
    function getTargetFormatDetails(formatValue) {
        switch (formatValue.toLowerCase()) {
            case 'png': return { mime: 'image/png', ext: '.png', qualityApplies: false, defaultQuality: 1.0 };
            case 'jpeg': return { mime: 'image/jpeg', ext: '.jpg', qualityApplies: true, defaultQuality: 0.92 };
            case 'webp': return { mime: 'image/webp', ext: '.webp', qualityApplies: true, defaultQuality: 0.80 };
            default: return null;
        }
    }

    function toggleConversionQualityUI() {
        const selectedFormat = convertFormatSelect.value;
        const details = getTargetFormatDetails(selectedFormat);
        if (details && details.qualityApplies) {
            convertQualityGroup.style.display = 'block';
            convertQualitySlider.value = details.defaultQuality * 100;
            convertQualityValue.textContent = convertQualitySlider.value;
        } else {
            convertQualityGroup.style.display = 'none';
        }
    }

    function generateConversionPreview() {
        convertPreviewOutput.innerHTML = '';
        convertApplyBtn.disabled = true;
        convertStatus.textContent = '';
        conversionPreviewData = [];

        const activeItems = getActiveItems();
        if (activeItems.length === 0) {
            convertStatus.textContent = 'No active images to convert.';
            convertStatus.style.color = 'var(--status-color)';
            return;
        }

        const targetFormatValue = convertFormatSelect.value;
        const targetDetails = getTargetFormatDetails(targetFormatValue);
        if (!targetDetails) {
            convertStatus.textContent = 'Invalid target format selected.';
            convertStatus.style.color = 'var(--error-color)';
            return;
        }
        const quality = targetDetails.qualityApplies ? parseInt(convertQualitySlider.value) / 100 : 1.0;

        const outputFragment = document.createDocumentFragment();
        let itemsToConvertCount = 0;

        activeItems.forEach(item => {
            const currentExt = item.originalFullImageName.substring(item.originalFullImageName.lastIndexOf('.')).toLowerCase();
            
            if (currentExt === targetDetails.ext) {
                return; // Already in target format
            }

            itemsToConvertCount++;
            const newImageFullName = item.imageName + targetDetails.ext; // Basename + new extension

            let newTagFullName = item.originalFullTagName; // Assume tag name doesn't change unless it's Kohya-style
            if (item.originalFullTagName) {
                // Check if original tag name was like "basename.original_ext.txt"
                const originalImageExtInTag = item.originalFullImageName.substring(item.originalFullImageName.lastIndexOf('.'));
                if (item.originalFullTagName === item.imageName + originalImageExtInTag + TAG_EXTENSION) {
                    newTagFullName = item.imageName + targetDetails.ext + TAG_EXTENSION;
                }
                // If it was just "basename.txt", newTagFullName will remain "basename.txt" (no change)
            }
            
            const oldDisplayPath = item.relativePath ? `${item.relativePath}/${item.originalFullImageName}` : item.originalFullImageName;
            const newDisplayPath = item.relativePath ? `${item.relativePath}/${newImageFullName}` : newImageFullName;
            
            let changeEntry = `<p><strong>${escapeHtml(oldDisplayPath)}</strong> → <strong>${escapeHtml(newDisplayPath)}</strong>`;
            if (newTagFullName && item.originalFullTagName !== newTagFullName) {
                 const oldTagDisplayPath = item.relativePath ? `${item.relativePath}/${item.originalFullTagName}` : item.originalFullTagName;
                 const newTagDisplayPath = item.relativePath ? `${item.relativePath}/${newTagFullName}` : newTagFullName;
                 changeEntry += `<br>&nbsp;&nbsp;&nbsp;<em>${escapeHtml(oldTagDisplayPath)}</em> → <em>${escapeHtml(newTagDisplayPath)}</em>`;
            } else if (newTagFullName && item.originalFullTagName === newTagFullName && item.tagHandle) {
                 changeEntry += `<br>&nbsp;&nbsp;&nbsp;<em>Tag file (${escapeHtml(newTagFullName)}) remains.</em>`;
            }
            changeEntry += `</p>`;
            const p = document.createElement('div');
            p.innerHTML = changeEntry;
            outputFragment.appendChild(p);

            conversionPreviewData.push({
                itemId: item.id,
                oldImageName: item.originalFullImageName,
                newImageName: newImageFullName,
                oldTagName: item.originalFullTagName,
                newTagName: newTagFullName,
                relativePath: item.relativePath,
                targetMime: targetDetails.mime,
                quality: quality
            });
        });

        if (itemsToConvertCount > 0) {
            convertPreviewOutput.appendChild(outputFragment);
            convertApplyBtn.disabled = false;
            convertStatus.textContent = `Previewing ${itemsToConvertCount} image(s) for conversion. Verify quality implications.`;
            convertStatus.style.color = 'var(--status-color)';
            if (targetDetails.mime === 'image/jpeg') {
                 convertStatus.textContent += " (PNG/GIF to JPEG may lose transparency/animation).";
            }
        } else {
            convertStatus.textContent = 'No images require conversion to the selected format.';
            convertStatus.style.color = 'var(--status-color)';
        }
    }
    
    async function convertAndSaveImageFile(item, targetMime, quality, newFullImageName, newFullTagName) {
        return new Promise(async (resolve, reject) => {
            const img = new Image();
            img.onload = async () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.naturalWidth;
                canvas.height = img.naturalHeight;
                const ctx = canvas.getContext('2d');
                if (targetMime === 'image/jpeg' || targetMime === 'image/webp') { // Fill background for formats that don't support transparency well by default
                    ctx.fillStyle = '#FFFFFF'; // White background
                    ctx.fillRect(0, 0, canvas.width, canvas.height);
                }
                ctx.drawImage(img, 0, 0);

                canvas.toBlob(async (blob) => {
                    if (!blob) {
                        reject(new Error(`Failed to create blob for ${item.originalFullImageName}`));
                        return;
                    }
                    try {
                        let parentDirHandle = datasetHandle;
                        if (item.relativePath) {
                            const parts = item.relativePath.split('/');
                            for (const p of parts) {
                                if (p) parentDirHandle = await parentDirHandle.getDirectoryHandle(p, { create: false });
                            }
                        }

                        // 1. Save new image file
                        const newImageFileHandle = await parentDirHandle.getFileHandle(newFullImageName, { create: true });
                        const writable = await newImageFileHandle.createWritable();
                        await writable.write(blob);
                        await writable.close();

                        // 2. Delete old image file (if different name)
                        if (item.originalFullImageName !== newFullImageName) {
                             await parentDirHandle.removeEntry(item.originalFullImageName);
                        }
                        
                        // Update item data
                        if (item.imageUrl) URL.revokeObjectURL(item.imageUrl);
                        const newFileForUrl = await newImageFileHandle.getFile();
                        item.imageUrl = URL.createObjectURL(newFileForUrl);
                        item.imageHandle = newImageFileHandle;
                        item.originalFullImageName = newFullImageName;
                        item.imageName = newFullImageName.substring(0, newFullImageName.lastIndexOf('.'));
                        item.fileSize = newFileForUrl.size;
                        item.lastModified = newFileForUrl.lastModified;


                        // 3. Rename tag file if necessary
                        if (item.tagHandle && item.originalFullTagName && newFullTagName && item.originalFullTagName !== newFullTagName) {
                            try {
                                await item.tagHandle.move(parentDirHandle, newFullTagName);
                                item.tagHandle = await parentDirHandle.getFileHandle(newFullTagName);
                                item.originalFullTagName = newFullTagName;
                            } catch (tagRenameError) {
                                console.warn(`Converted image but failed to rename tag file for ${item.id}: ${tagRenameError.message}`);
                                // Continue, as image conversion was the primary goal
                            }
                        } else if (item.tagHandle && item.originalFullTagName && !newFullTagName) {
                            // This case should ideally not happen if logic in preview is correct
                            // Means a tag file existed but new logic says it shouldn't have a new name (e.g. it was image.txt and base didn't change)
                            // No action needed for tag file name itself.
                        }
                        resolve();
                    } catch (fileError) {
                        reject(fileError);
                    }
                }, targetMime, quality);
            };
            img.onerror = () => {
                reject(new Error(`Failed to load image ${item.originalFullImageName} for conversion.`));
            };
            img.src = item.imageUrl; // Trigger load
        });
    }


    async function applyImageConversion() {
        if (conversionPreviewData.length === 0 || convertApplyBtn.disabled) {
            convertStatus.textContent = 'No conversions to apply or preview first.';
            convertStatus.style.color = 'var(--error-color)';
            return;
        }
        if (!confirm(`This will convert ${conversionPreviewData.length} image(s) on your disk. Original files will be replaced. This action is IRREVERSIBLE and may affect image quality. Proceed?`)) {
            convertStatus.textContent = 'Conversion cancelled by user.';
            convertStatus.style.color = 'var(--status-color)';
            return;
        }

        updateStatus('Applying image format conversions...', true);
        convertApplyBtn.disabled = true;
        let successCount = 0;
        let errorCount = 0;

        try {
            if (!datasetHandle || (await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted' &&
                await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted')) {
                updateStatus('Error: Write permission denied.');
                convertStatus.textContent = 'Write permission denied.';
                convertStatus.style.color = 'var(--error-color)';
                return;
            }

            for (const conv of conversionPreviewData) {
                const item = allImageData.find(i => i.id === conv.itemId);
                if (!item || item.isTrashed) {
                    console.warn(`Skipping conversion for item ${conv.itemId} (not found or trashed).`);
                    continue;
                }
                try {
                    await convertAndSaveImageFile(item, conv.targetMime, conv.quality, conv.newImageName, conv.newTagName);
                    successCount++;
                } catch (conversionError) {
                    console.error(`Error converting item ${item.id} (${item.originalFullImageName}): ${conversionError.message}`);
                    errorCount++;
                }
            }
        } catch (e) {
            updateStatus(`Error during conversion process: ${e.message}`, false);
            convertStatus.textContent = `Error: ${e.message}`;
            convertStatus.style.color = 'var(--error-color)';
        } finally {
            conversionPreviewData = [];
            convertPreviewOutput.innerHTML = '';
            convertApplyBtn.disabled = true;

            let finalMsg = `Conversion complete. ${successCount} image(s) converted.`;
            if (errorCount > 0) finalMsg += ` ${errorCount} errors (see console).`;
            convertStatus.textContent = finalMsg;
            convertStatus.style.color = errorCount > 0 ? 'var(--error-color)' : 'var(--success-color)';

            updateStatus(finalMsg, false);
            filterAndSearch(); // Re-render grid with new image URLs and info
            recalculateAllTagStats(); // Though tags themselves didn't change, file info did
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
                event.preventDefault(); updateActiveSuggestion((activeSuggestionIndex + 1) % currentSuggestions.length); return;
            }
            if (event.key === 'ArrowUp') {
                event.preventDefault(); updateActiveSuggestion((activeSuggestionIndex - 1 + currentSuggestions.length) % currentSuggestions.length); return;
            }
            if ((event.key === 'Enter' || event.key === 'Tab') && activeSuggestionIndex !== -1) {
                event.preventDefault(); selectSuggestion(currentSuggestions[activeSuggestionIndex]); return;
            }
            if (event.key === 'Escape') { event.preventDefault(); hideSuggestions(); return; }
        }

        if (event.key === 'Escape') {
            if (isImageZoomed) { toggleEditorImageZoom(); return; }
            if (isRuleEditorOpen) { closeRuleEditor(); return; }
            if (isSingleEditorOpen) { closeEditor(); return; }
            if (isBatchEditorOpen) { closeBatchEditor(); return; }
            if (selectedItemIds.size > 0) { clearSelection(); return; }
            if (isInputFocusedGeneral && activeElement.value !== '') {
                activeElement.value = '';
                if (activeElement.id === 'tag-frequency-search') activeElement.dispatchEvent(new Event('input', { bubbles: true }));
                if (activeElement.id === 'unify-preview-output') { unifyPreviewOutput.innerHTML = ''; unifyApplyBtn.disabled = true; unifyStatus.textContent = ''; unificationPreviewData = []; }
                if (activeElement.id === 'convert-preview-output') { convertPreviewOutput.innerHTML = ''; convertApplyBtn.disabled = true; convertStatus.textContent = ''; conversionPreviewData = []; }
                return;
            }
            if (keyboardFocusIndex !== -1) { clearKeyboardFocus(); return; }
        }

        if (event.key === 'Enter' && isInputFocusedGeneral && !isModalOpen) {
            let buttonToClick = null;
            switch (activeElement.id) {
                case 'search-tags': case 'exclude-tags': buttonToClick = filterBtn; break;
                case 'rename-old-tag': case 'rename-new-tag': buttonToClick = renameTagBtn; break;
                case 'remove-tag-name': buttonToClick = removeTagBtn; break;
                case 'add-trigger-words': buttonToClick = addTagsBtn; break;
                case 'tag-frequency-search': event.preventDefault(); return; 
                case 'sort-property': case 'sort-order': buttonToClick = applySortBtn; break;
                case 'unify-prefix': case 'unify-start-number': case 'unify-padding-length': buttonToClick = unifyPreviewBtn; break;
                case 'convert-format-select': case 'convert-quality-slider': buttonToClick = convertPreviewBtn; break;
            }
            if (buttonToClick && !buttonToClick.disabled) { event.preventDefault(); buttonToClick.click(); return; }
        }
        
        const currentEditorItem = (currentEditIndex !== -1) ? allImageData[currentEditIndex] : null;
        if (isSingleEditorOpen && !isTextareaFocused && !isImageZoomed && (!currentEditorItem || !currentEditorItem.isTrashed)) { 
            if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
                event.preventDefault();
                navigateEditor(event.key === 'ArrowLeft' ? 'previous' : 'next');
                return;
            }
        }

        if (!isModalOpen && !isInputFocusedGeneral && displayedImageDataIndices.length > 0) {
            let newFocusDisplayIndex = keyboardFocusIndex; 
            let handled = false;
            if (newFocusDisplayIndex === -1 && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
                newFocusDisplayIndex = 0; handled = true;
            } else if (newFocusDisplayIndex !== -1) { 
                let columns = 1;
                const gridWidth = imageGrid.offsetWidth;
                const firstItemElement = imageGrid.querySelector('.grid-item:not(.trashed-item)'); 
                if (firstItemElement) {
                    const itemStyle = getComputedStyle(firstItemElement);
                    const itemOuterWidth = firstItemElement.offsetWidth + parseInt(itemStyle.marginLeft) + parseInt(itemStyle.marginRight);
                    if (itemOuterWidth > 0) columns = Math.max(1, Math.floor(gridWidth / itemOuterWidth));
                }
                switch (event.key) {
                    case 'ArrowLeft': newFocusDisplayIndex = Math.max(0, newFocusDisplayIndex - 1); handled = true; break;
                    case 'ArrowRight': newFocusDisplayIndex = Math.min(displayedImageDataIndices.length - 1, newFocusDisplayIndex + 1); handled = true; break;
                    case 'ArrowUp': newFocusDisplayIndex = Math.max(0, newFocusDisplayIndex - columns); handled = true; break;
                    case 'ArrowDown': newFocusDisplayIndex = Math.min(displayedImageDataIndices.length - 1, newFocusDisplayIndex + columns); handled = true; break;
                    case 'Enter':
                        const globalIdx = displayedImageDataIndices[keyboardFocusIndex];
                        const itemToOpen = allImageData[globalIdx];
                        if (itemToOpen && !itemToOpen.isTrashed) { 
                            clearSelection(); addSelection(itemToOpen.id); openEditor(itemToOpen.id);
                        }
                        handled = true; break;
                }
            }
            if (handled) {
                event.preventDefault();
                if (newFocusDisplayIndex !== keyboardFocusIndex || event.key === 'Enter') {
                    setKeyboardFocus(newFocusDisplayIndex);
                }
            }
        }
    }

    function toggleEditorImageZoom() {
        editorImagePreview.classList.toggle('zoomed');
        editorModalContent.classList.toggle('image-zoomed'); 
    }

    function handleDocumentClickToClearSelection(event) {
        const isModalOpen = !editorModal.classList.contains('hidden') || 
                            !batchEditorModal.classList.contains('hidden') || 
                            !ruleEditorModal.classList.contains('hidden');
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
    emptyTrashBtn.addEventListener('click', permanentlyDeleteTrashedItems);
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
    
    findDuplicateFilesBtn.addEventListener('click', findDuplicateTagFiles);
    findOrphanedMissingBtn.addEventListener('click', findOrphanedAndMissingTagFiles);
    findCaseInconsistentTagsBtn.addEventListener('click', findCaseInconsistentTags);
    exportUniqueTagsBtn.addEventListener('click', exportUniqueTagList);

    unifyPadZerosCheckbox.addEventListener('change', toggleUnifyPaddingLengthGroup);
    unifyPreviewBtn.addEventListener('click', generateUnificationPreview);
    unifyApplyBtn.addEventListener('click', applyUnification);

    convertFormatSelect.addEventListener('change', toggleConversionQualityUI);
    convertQualitySlider.addEventListener('input', () => { convertQualityValue.textContent = convertQualitySlider.value; });
    convertPreviewBtn.addEventListener('click', generateConversionPreview);
    convertApplyBtn.addEventListener('click', applyImageConversion);


    createNewRuleBtn.addEventListener('click', () => openRuleEditor());
    ruleEditorCancelBtn.addEventListener('click', closeRuleEditor);
    ruleEditorModal.addEventListener('click', (e) => { if (e.target === ruleEditorModal) closeRuleEditor(); });
    bulkApplyCustomRulesBtn.addEventListener('click', performBulkCustomRulesOperation);

    tagFrequencySearchInput.addEventListener('input', () => {
        clearTimeout(debouncedUpdateFrequencyList);
        debouncedUpdateFrequencyList = setTimeout(updateStatisticsDisplay, 300);
    });

    editorSaveBtn.addEventListener('click', () => { saveEditorChanges(); hideSuggestions(); });
    editorCancelBtn.addEventListener('click', closeEditor);
    editorRemoveDupsBtn.addEventListener('click', () => { removeDuplicateTagsInEditor(); hideSuggestions(); });
    editorModal.addEventListener('click', (e) => { if (e.target === editorModal && !editorImagePreview.classList.contains('zoomed')) closeEditor(); });
    
    batchEditorApplyBtn.addEventListener('click', applyBatchEdits);
    batchEditorMoveToTrashBtn.addEventListener('click', batchMoveSelectedToTrash);
    batchEditorCancelBtn.addEventListener('click', closeBatchEditor);
    batchEditorModal.addEventListener('click', (e) => { if (e.target === batchEditorModal) closeBatchEditor(); });

    editorTagsTextarea.addEventListener('input', (e) => { showSuggestions(e.target); });
    editorTagsTextarea.addEventListener('blur', () => { setTimeout(() => { if (!editorTagSuggestionsBox.contains(document.activeElement)) hideSuggestions(); }, 150); });
    editorTagsTextarea.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); saveEditorChanges(); hideSuggestions(); } });
    
    editorImagePreview.addEventListener('click', toggleEditorImageZoom);
    if (editorPrevBtn) editorPrevBtn.addEventListener('click', () => navigateEditor('previous'));
    if (editorNextBtn) editorNextBtn.addEventListener('click', () => navigateEditor('next'));

    document.addEventListener('keydown', handleGlobalKeyDown);
    document.addEventListener('click', handleDocumentClickToClearSelection);

    // --- Initial Setup ---
    loadInitialTheme();
    loadCustomRules(); 
    toggleUnifyPaddingLengthGroup(); 
    toggleConversionQualityUI(); // Initial state for conversion quality UI
    sortPropertySelect.value = currentSortProperty;
    sortOrderSelect.value = currentSortOrder;
    updateStatus('Ready. Load a dataset folder.');
    updateSaveAllButtonState();
    updateEmptyTrashButtonState();

    if (!window.showDirectoryPicker) {
        updateStatus("Warning: Browser lacks File System Access API. Core functionality limited.");
        [loadDatasetBtn, saveAllBtn, emptyTrashBtn, batchEditSelectionBtn, applySortBtn, createNewRuleBtn, bulkApplyCustomRulesBtn,
         renameTagBtn, removeTagBtn, addTagsBtn, removeDuplicatesBtn, 
         findDuplicateFilesBtn, findOrphanedMissingBtn, findCaseInconsistentTagsBtn, exportUniqueTagsBtn,
         unifyPreviewBtn, unifyApplyBtn, convertPreviewBtn, convertApplyBtn]
        .forEach(btn => { if (btn) btn.disabled = true; });
    }
});