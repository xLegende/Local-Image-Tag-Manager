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
    let allFileSystemEntries = new Map();
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
                            imageName: imageNameForRecord, 
                            originalFullImageName: pair.originalImageName, 
                            originalFullTagName: pair.originalTagName, 
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
                            isTrashed: false, // New property for trash system
                        });
                        // Initial tag stats update only considers non-trashed items, so this is fine
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
            sortAndRender(); // This will internally use getActiveItems for display
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
        updateStatisticsDisplay(); // Will show 0s
        updateDisplayedCount();
        updateSaveAllButtonState();
        updateEmptyTrashButtonState();
        duplicateFilesOutput.innerHTML = '';
        orphanedMissingOutput.innerHTML = '';
        caseInconsistentOutput.innerHTML = '';
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
            if (!item) return; // Should not happen if displayedImageDataIndices is correct

            const div = document.createElement('div');
            div.classList.add('grid-item');
            div.dataset.globalIndex = globalIndex; // Store global index for easier lookup
            div.dataset.displayIndex = displayIndex; // Store current display index
            div.id = item.id;

            if (item.modified) div.classList.add('modified');
            if (item.isTrashed) div.classList.add('trashed-item');
            if (selectedItemIds.has(item.id)) div.classList.add('selected');

            div.tabIndex = -1; // For focus management, though actual focus handled by keyboard-focus class
            div.style.outline = 'none'; // Remove default focus outline

            div.addEventListener('click', (e) => handleGridItemInteraction(e, item.id, globalIndex, displayIndex));

            // Action button (Move to Trash / Restore)
            const actionBtn = document.createElement('span');
            actionBtn.classList.add('image-action-btn');
            if (item.isTrashed) {
                actionBtn.classList.add('restore-btn');
                actionBtn.textContent = '♻'; // Restore icon (or text)
                actionBtn.title = `Restore "${item.originalFullImageName}" from trash`;
                actionBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    restoreFromTrash(item.id);
                });
            } else {
                actionBtn.classList.add('move-to-trash-btn');
                actionBtn.textContent = '×'; // Delete icon
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
            renderTagsForItem(tagsDiv, item); // Pass item itself
            div.appendChild(tagsDiv);

            fragment.appendChild(div);
        });
        imageGrid.appendChild(fragment);
        updateDisplayedCount();
        updateBatchEditButtonState();
    }


    function handleGridItemInteraction(event, itemId, globalItemIndex, displayItemIndex) {
        const item = allImageData[globalItemIndex];
        if (!item) return; // Should not happen

        // Clicking on the image of a non-trashed item, if it's the only selected item, opens editor
        const isImageClick = event.target.tagName === 'IMG';
        const canOpenEditor = !item.isTrashed && isImageClick;

        if (event.shiftKey && lastInteractedItemId && lastInteractedItemId !== itemId) {
            // Find display indices for range selection
            const lastInteractedGlobalIndex = findIndexById(lastInteractedItemId);
            const lastInteractedDisplayIndex = displayedImageDataIndices.indexOf(lastInteractedGlobalIndex);
            const currentDisplayIndex = displayItemIndex;

            if (lastInteractedDisplayIndex === -1 || currentDisplayIndex === -1) { // One of them isn't in current display
                const wasSelected = selectedItemIds.has(itemId);
                const wasOnlySelection = wasSelected && selectedItemIds.size === 1;
                if(wasOnlySelection) {
                    selectedItemIds.delete(itemId);
                    document.getElementById(itemId)?.classList.remove('selected');
                    lastInteractedItemId = null;
                } else {
                    clearSelection();
                    if (!item.isTrashed) addSelection(itemId); // Only select non-trashed
                    lastInteractedItemId = item.isTrashed ? null : itemId;
                }
            } else {
                const start = Math.min(lastInteractedDisplayIndex, currentDisplayIndex);
                const end = Math.max(lastInteractedDisplayIndex, currentDisplayIndex);

                if (!(event.ctrlKey || event.metaKey)) { // If not holding Ctrl/Cmd, clear previous selection
                    clearSelection();
                }
                for (let i = start; i <= end; i++) {
                    const gIdx = displayedImageDataIndices[i];
                    if (allImageData[gIdx] && !allImageData[gIdx].isTrashed) { // Only select non-trashed
                        addSelection(allImageData[gIdx].id);
                    }
                }
            }
        } else if (event.ctrlKey || event.metaKey) {
            if (!item.isTrashed) toggleSelection(itemId); // Only select non-trashed
            lastInteractedItemId = item.isTrashed ? null : itemId;
        } else { // Simple click
            const wasSelected = selectedItemIds.has(itemId);
            const wasOnlySelection = wasSelected && selectedItemIds.size === 1;

            if (wasOnlySelection) { // Clicked on the only selected item
                selectedItemIds.delete(itemId);
                document.getElementById(itemId)?.classList.remove('selected');
                lastInteractedItemId = null;
                // Do not open editor if deselecting
            } else {
                clearSelection();
                if (!item.isTrashed) {
                    addSelection(itemId);
                    lastInteractedItemId = itemId;
                    if (canOpenEditor) { // Open editor only if it's a fresh, single selection via image click
                         openEditor(itemId);
                    }
                } else {
                    lastInteractedItemId = null; // Cannot select trashed item
                }
            }
        }
        updateBatchEditButtonState();
    }

    function toggleSelection(itemId) {
        const element = document.getElementById(itemId);
        const item = allImageData[findIndexById(itemId)];
        if (item && item.isTrashed) return; // Do not select trashed items

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
        if (item && item.isTrashed) return; // Do not select trashed items

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

    function renderTagsForItem(tagsContainer, item) { // item is the full data object
        tagsContainer.innerHTML = '';
        if (item.tags.length === 0) {
            tagsContainer.textContent = '(no tags)';
        } else {
            item.tags.forEach(tag => {
                const tagSpan = document.createElement('span');
                tagSpan.classList.add('tag');
                tagSpan.textContent = tag;

                if (!item.isTrashed) { // Only add delete button for non-trashed items
                    const deleteBtn = document.createElement('span');
                    deleteBtn.classList.add('tag-delete-btn');
                    deleteBtn.textContent = '×';
                    deleteBtn.title = `Remove tag "${tag}"`;
                    deleteBtn.addEventListener('click', (e) => {
                        e.stopPropagation(); // Prevent grid item click
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
        displayedCountSpan.textContent = displayedImageDataIndices.length; // Number of items currently in the grid
        totalCountSpan.textContent = activeItemsCount; // Total non-trashed items in the dataset
    }

    function recalculateAllTagStats() { // Operates on active items
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

    function updateTagStatsIncremental(addedTags, removedTags) { // Assumes called for an active item
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
        // No need to call updateStatisticsDisplay here, usually called by parent
    }

    function updateTagStats(addedTags, removedTags, isInitialLoad = false) {
        // For initial load, we build up from scratch. For incremental, we adjust.
        if (isInitialLoad) { // During dataset load, items are not yet trashed
            addedTags.map(t => t.trim()).filter(t => t).forEach(tag => {
                tagFrequencies.set(tag, (tagFrequencies.get(tag) || 0) + 1);
                uniqueTags.add(tag);
            });
        } else { // Incremental update (e.g., after tag edit) - assumes item is active
            updateTagStatsIncremental(addedTags, removedTags);
        }
        // updateStatisticsDisplay will be called by the caller (e.g., loadDataset, saveEditorChanges)
    }


    function updateStatisticsDisplay() {
        const activeItems = getActiveItems();
        statsTotalImages.textContent = activeItems.length;
        statsUniqueTags.textContent = uniqueTags.size; // uniqueTags is already based on active items via recalculateAllTagStats
        statsModifiedImages.textContent = activeItems.filter(item => item.modified).length;

        tagFrequencyList.innerHTML = '';
        const searchTerm = tagFrequencySearchInput.value.trim().toLowerCase();
        // tagFrequencies should already be based on active items
        const filteredFrequencies = [...tagFrequencies.entries()].filter(([t]) => !searchTerm || t.toLowerCase().includes(searchTerm));
        const sortedTags = filteredFrequencies.sort((a, b) => (b[1] - a[1]) || a[0].localeCompare(b[0]));

        if (sortedTags.length === 0) {
            tagFrequencyList.innerHTML = `<div><span>${searchTerm ? 'No active tags match search.' : 'No active tags loaded.'}</span></div>`;
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
        updateDisplayedCount(); // Ensure grid header count is also up-to-date
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
        // Sort displayedImageDataIndices, which should already point to correct items from allImageData
        displayedImageDataIndices.sort((indexA, indexB) => {
            const itemA = allImageData[indexA];
            const itemB = allImageData[indexB];
            let comparison = 0;
            if (!itemA || !itemB) return 0; // Should not happen with valid indices

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
            // Determine the base pool of items to filter from
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
                // Apply specific non-trash filters only if not in "trashed" mode
                if (filterMode !== 'trashed') {
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
                        // 'all' (active) means no further filtering here based on these criteria
                    }
                }
                if (!include) return false;

                // Apply search term filters
                if (searchTermsRaw.length > 0) {
                    const itemTagsLower = item.tags.map(t => t.toLowerCase());
                    if (searchMode === 'AND') {
                        include = searchTermsRaw.every(term => itemTagsLower.some(tag => tag.includes(term.toLowerCase())));
                    } else { // OR mode
                        include = searchTermsRaw.some(term => itemTagsLower.some(tag => tag.includes(term.toLowerCase())));
                    }
                }
                if (!include) return false;

                // Apply exclude term filters
                if (excludeTerms.length > 0) {
                    include = !excludeTerms.some(term => item.tags.includes(term));
                }
                return include;
            });
            sortAndRender(); // This will render the filtered and sorted displayedImageDataIndices
        });
    }

    function resetFilters() {
        clearSelection();
        searchTagsInput.value = '';
        excludeTagsInput.value = '';
        searchModeSelect.value = 'AND';
        filterUntaggedSelect.value = 'all'; // Default to 'all active'
        duplicateFilesOutput.innerHTML = '';
        orphanedMissingOutput.innerHTML = '';
        caseInconsistentOutput.innerHTML = '';
        currentSortProperty = 'path';
        sortPropertySelect.value = currentSortProperty;
        currentSortOrder = 'asc';
        sortOrderSelect.value = currentSortOrder;
        
        // Reset displayedImageDataIndices to all active items
        displayedImageDataIndices = getActiveItems().map(item => allImageData.indexOf(item));
        sortAndRender();
    }


    function findIndexById(itemId) {
        return allImageData.findIndex(item => item.id === itemId);
    }

    function openEditor(itemId) {
        const index = findIndexById(itemId); // index in allImageData
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
        // editorStatus already handled above for trashed items
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
        // Restore focus to grid item if applicable
        if (keyboardFocusIndex !== -1 && displayedImageDataIndices.length > keyboardFocusIndex) {
            const globalFocusedIndex = displayedImageDataIndices[keyboardFocusIndex];
            if (allImageData[globalFocusedIndex]) {
                 const focusedItemId = allImageData[globalFocusedIndex].id;
                 document.getElementById(focusedItemId)?.focus(); // Re-focus grid item
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
            const oldTagsForStatUpdate = parseTags(currentEditOriginalTagsString); // Tags before this edit session
            
            item.tags = newTagsFromTextarea;
            item.modified = true;
            
            const addedForStatUpdate = newTagsFromTextarea.filter(t => !oldTagsForStatUpdate.includes(t));
            const removedForStatUpdate = oldTagsForStatUpdate.filter(t => !newTagsFromTextarea.includes(t));
            
            updateTagStatsIncremental(addedForStatUpdate, removedForStatUpdate); // Update stats based on delta
            recalculateAllTagStats(); // Recalculate global unique tags and frequencies
            
            const gridItem = document.getElementById(item.id);
            if (gridItem) {
                gridItem.classList.add('modified');
                const tagsDiv = gridItem.querySelector('.tags');
                if (tagsDiv) renderTagsForItem(tagsDiv, item);
            }
            editorStatus.textContent = 'Tags updated. Remember to "Save All Changes".';
            editorStatus.style.color = 'var(--success-color)';
            currentEditOriginalTagsString = newTagsStringFromTextarea; // Update original for next comparison
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
        if (item.isTrashed) return; // Cannot modify trashed items

        const oldLen = item.tags.length;
        const newTags = item.tags.filter(t => t !== tagToRemove);

        if (newTags.length < oldLen) {
            item.tags = newTags;
            item.modified = true;
            updateSaveAllButtonState();
            updateTagStatsIncremental([], [tagToRemove]); // Update stats
            recalculateAllTagStats(); // Recalculate global unique tags and frequencies

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
        if (item.isTrashed) return; // Already trashed

        item.isTrashed = true;
        item.modified = false; // Trashing is not a "tag modification" for save all button

        // Update UI for the specific item
        const gridItemElement = document.getElementById(item.id);
        if (gridItemElement) {
            gridItemElement.classList.add('trashed-item');
            gridItemElement.classList.remove('modified'); // Visually remove modified if it was, as it's now just "trashed"
            const actionBtn = gridItemElement.querySelector('.image-action-btn');
            if (actionBtn) {
                actionBtn.classList.remove('move-to-trash-btn');
                actionBtn.classList.add('restore-btn');
                actionBtn.textContent = '♻';
                actionBtn.title = `Restore "${item.originalFullImageName}" from trash`;
                actionBtn.replaceWith(actionBtn.cloneNode(true)); // Re-attach to clear old listeners
                gridItemElement.querySelector('.restore-btn').addEventListener('click', (e) => {
                    e.stopPropagation();
                    restoreFromTrash(item.id);
                });
            }
            const tagsDiv = gridItemElement.querySelector('.tags');
            if (tagsDiv) renderTagsForItem(tagsDiv, item); // Re-render tags to remove delete buttons
        }
        
        // If item was selected, deselect it as trashed items cannot be part of batch edits
        if (selectedItemIds.has(item.id)) {
            selectedItemIds.delete(item.id);
            gridItemElement?.classList.remove('selected');
            updateBatchEditButtonState();
        }
        if (lastInteractedItemId === item.id) lastInteractedItemId = null;

        // Recalculate stats and update counts
        recalculateAllTagStats(); // This will exclude the newly trashed item
        updateEmptyTrashButtonState();
        updateSaveAllButtonState(); // Ensure save button reflects actual tag modifications

        // If current view is not "Only Trashed", the item might disappear. Re-filter.
        if (filterUntaggedSelect.value !== 'trashed') {
            filterAndSearch(); // Re-apply filters, which will hide the item if not showing trash
        } else {
             // If current view IS "Only Trashed", we still need to re-render the grid
             // to show the item now correctly styled as trashed, or if sorting changed.
            sortAndRender();
        }
        updateStatus(`Moved "${item.originalFullImageName}" to trash.`);
    }

    function restoreFromTrash(itemId) {
        const index = findIndexById(itemId);
        if (index === -1) return;
        const item = allImageData[index];
        if (!item.isTrashed) return; // Not in trash

        item.isTrashed = false;
        // item.modified remains as it was before trashing (if it was modified for tags)

        const gridItemElement = document.getElementById(item.id);
        if (gridItemElement) {
            gridItemElement.classList.remove('trashed-item');
            if(item.modified) gridItemElement.classList.add('modified'); // Re-add if it was modified

            const actionBtn = gridItemElement.querySelector('.image-action-btn');
            if (actionBtn) {
                actionBtn.classList.remove('restore-btn');
                actionBtn.classList.add('move-to-trash-btn');
                actionBtn.textContent = '×';
                const deleteTitle = item.relativePath ? `Move "${item.relativePath}/${item.originalFullImageName}" to trash` : `Move "${item.originalFullImageName}" to trash`;
                actionBtn.title = deleteTitle;
                actionBtn.replaceWith(actionBtn.cloneNode(true)); // Re-attach
                gridItemElement.querySelector('.move-to-trash-btn').addEventListener('click', (e) => {
                    e.stopPropagation();
                    promptMoveToTrash(item.id);
                });
            }
             const tagsDiv = gridItemElement.querySelector('.tags');
            if (tagsDiv) renderTagsForItem(tagsDiv, item); // Re-render tags to add delete buttons
        }

        recalculateAllTagStats(); // Will now include this item
        updateEmptyTrashButtonState();
        updateSaveAllButtonState();

        // If current view is "Only Trashed", the item will disappear. Re-filter.
        if (filterUntaggedSelect.value === 'trashed') {
            filterAndSearch();
        } else {
            // If current view is NOT "Only Trashed", we still need to re-render the grid
            // to show the item now correctly styled as active, or if sorting changed.
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
                // Delete image file
                if (item.imageHandle) {
                    try {
                        let parentDirHandle = datasetHandle;
                        if (item.relativePath) {
                            const parts = item.relativePath.split('/');
                            for (const p of parts) {
                                if (p) parentDirHandle = await parentDirHandle.getDirectoryHandle(p, { create: false });
                            }
                        }
                        await parentDirHandle.removeEntry(item.imageHandle.name);
                    } catch (e) {
                        console.error(`Error deleting image file ${item.originalFullImageName}:`, e);
                        itemError = true;
                    }
                }
                // Delete tag file
                if (item.tagHandle) {
                     try {
                        let parentDirHandle = datasetHandle;
                        if (item.relativePath) {
                            const parts = item.relativePath.split('/');
                            for (const p of parts) {
                                if (p) parentDirHandle = await parentDirHandle.getDirectoryHandle(p, { create: false });
                            }
                        }
                        await parentDirHandle.removeEntry(item.tagHandle.name);
                    } catch (e) {
                        console.error(`Error deleting tag file ${item.originalFullTagName}:`, e);
                        // Don't set itemError to true if image already failed, to avoid double counting error
                        if (!itemError) itemError = true;
                    }
                }
                if (item.imageUrl) URL.revokeObjectURL(item.imageUrl);
                if (itemError) errorCount++; else deletedCount++;
            }
        } catch (permError) {
            updateStatus(`Permission error during deletion: ${permError.message}`);
            // No full stop, try to update state as much as possible
        } finally {
            // Update allImageData by removing all items marked isTrashed, regardless of FS operation success
            // This ensures they are gone from the UI's perspective.
            const newAllImageData = allImageData.filter(item => !item.isTrashed);
            //const actuallyRemovedCount = allImageData.length - newAllImageData.length; // Should match itemsToDelete.length
            allImageData = newAllImageData;
            
            clearSelection(); // Clear any selections as items are gone
            clearKeyboardFocus();

            // Recalculate everything
            recalculateAllTagStats();
            updateEmptyTrashButtonState(); // Will go to 0
            updateSaveAllButtonState();

            // Refresh the grid based on current filters
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
        activeItems.forEach((item) => { // Iterate only over active items
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
                        currentTagsArray = [...uniqueTagsToAdd, ...currentTagsArray]; // Add to beginning
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
            // Check if tags actually changed content-wise, not just order for non-duplicate ops
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
                const itemData = allImageData.find(i => i.id === id); // Get from allImageData for rendering
                if (gridItem && itemData) {
                    gridItem.classList.add('modified');
                    const tagsDiv = gridItem.querySelector('.tags');
                    if (tagsDiv) renderTagsForItem(tagsDiv, itemData);
                }
            });
            if (statsChanged) recalculateAllTagStats(); // Recalculate based on all active items
            updateStatus(`${modCount} active items affected by "${desc}". Save changes.`, false);
        } else {
            updateStatus(`Bulk op "${desc}" complete. No changes to active items.`, false);
        }
        // Clear inputs
        if (operation === 'rename') { renameOldTagInput.value = ''; renameNewTagInput.value = ''; }
        if (operation === 'remove') removeTagNameInput.value = '';
        if (operation === 'add') addTriggerWordsInput.value = '';
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

    async function findOrphanedAndMissingTagFiles() { // This checks the entire filesystem entries vs allImageData
        if (!datasetHandle) {
            orphanedMissingOutput.innerHTML = '<p>Load dataset first.</p>'; return;
        }
        updateStatus('Checking for orphaned/missing tag files...', true);
        const orphanedTagFiles = []; // Tag files on disk with no corresponding image in allImageData
        const imagesMissingTagFiles = []; // Images in allImageData that don't have a tagHandle
        const allLoadedImageBasePaths = new Set(); // Stores base_path/image_name (lowercase) for all images in allImageData
        const allLoadedTagBasePaths = new Set();   // Stores base_path/tag_name (lowercase) for all tags in allImageData

        allImageData.forEach(item => {
            const imgBase = item.imageName.toLowerCase();
            const imgFullPathKey = item.relativePath ? `${item.relativePath}/${imgBase}` : imgBase;
            allLoadedImageBasePaths.add(imgFullPathKey);
            if (item.tagHandle) {
                const tagBase = item.originalFullTagName.substring(0, item.originalFullTagName.lastIndexOf('.')).toLowerCase();
                const tagFullPathKey = item.relativePath ? `${item.relativePath}/${tagBase}` : tagBase;
                allLoadedTagBasePaths.add(tagFullPathKey);
            } else if (!item.isTrashed) { // If it's an active image and has no tagHandle, it's missing a tag file
                 imagesMissingTagFiles.push(item.relativePath ? `${item.relativePath}/${item.originalFullImageName}` : item.originalFullImageName);
            }
        });
        
        for (const [fullPathOnDisk, entryInfo] of allFileSystemEntries.entries()) {
            if (entryInfo.kind !== 'file') continue;

            const lowerCaseName = entryInfo.name.toLowerCase();
            const ext = lowerCaseName.substring(lowerCaseName.lastIndexOf('.'));
            const baseName = lowerCaseName.substring(0, lowerCaseName.lastIndexOf('.'));
            const pathParts = fullPathOnDisk.split('/'); pathParts.pop(); 
            const relativePath = pathParts.join('/');
            const diskFileKey = relativePath ? `${relativePath}/${baseName}` : baseName;

            if (ext === TAG_EXTENSION) { // Found a .txt file on disk
                if (!allLoadedImageBasePaths.has(diskFileKey)) { // No corresponding image loaded in our app
                    orphanedTagFiles.push(fullPathOnDisk);
                }
            }
        }

        orphanedMissingOutput.innerHTML = '';
        let foundAny = false;
        if (imagesMissingTagFiles.length > 0) {
            foundAny = true;
            const p = document.createElement('p');
            p.textContent = `Active images missing tag files (${imagesMissingTagFiles.length}):`;
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
            p.textContent = `Orphaned tag files on disk (no matching image loaded) (${orphanedTagFiles.length}):`;
            orphanedMissingOutput.appendChild(p);
            const ul = document.createElement('ul');
            orphanedTagFiles.sort().forEach(filePath => {
                const li = document.createElement('li'); li.textContent = escapeHtml(filePath); ul.appendChild(li);
            });
            orphanedMissingOutput.appendChild(ul);
        }
        if (!foundAny) {
            orphanedMissingOutput.innerHTML = '<p>No orphaned or missing tag files found according to current data load.</p>';
        }
        updateStatus('Orphaned/missing tag file check complete.', false);
    }

    function findCaseInconsistentTags() { // Uses uniqueTags, which is based on active items
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

    function exportUniqueTagList() { // Uses uniqueTags, based on active items
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


    async function saveAllChanges() { // Saves tag modifications for active items
        if (!datasetHandle) {
            updateStatus("No dataset loaded."); return;
        }
        const itemsToSave = getActiveItems().filter(item => item.modified);
        if (itemsToSave.length === 0) {
            updateStatus("No tag modifications on active items to save.");
            updateSaveAllButtonState(); // Should disable it
            return;
        }
        let savedCount = 0, errorCount = 0;
        updateStatus(`Saving ${itemsToSave.length} modified tag file(s) for active items...`, true);
        try {
            if (await datasetHandle.queryPermission({ mode: 'readwrite' }) !== 'granted' &&
                await datasetHandle.requestPermission({ mode: 'readwrite' }) !== 'granted') {
                updateStatus('Error: Write permission denied. Cannot save tags.');
                return; // Do not disable save button here, user might retry
            }
        } catch (permError) {
            updateStatus(`Error requesting permission: ${permError.message}`);
            return;
        }

        for (const item of itemsToSave) { // item is already confirmed active and modified
            const newTagString = joinTags(item.tags);
            const tagFileName = item.originalFullTagName || `${item.imageName}${TAG_EXTENSION}`;
            const fullLogPath = item.relativePath ? `${item.relativePath}/${tagFileName}` : tagFileName;
            try {
                let tagHandle = item.tagHandle;
                if (!tagHandle) { // Need to create the tag file
                    try {
                        let parentDirHandle = datasetHandle;
                        if (item.relativePath) {
                            const pathParts = item.relativePath.split('/');
                            for (const part of pathParts) {
                                if (part) parentDirHandle = await parentDirHandle.getDirectoryHandle(part, { create: false });
                            }
                        }
                        tagHandle = await parentDirHandle.getFileHandle(tagFileName, { create: true });
                        item.tagHandle = tagHandle;
                        item.originalFullTagName = tagFileName;
                    } catch (handleError) {
                        console.error(`Error getting/creating tag file handle for ${fullLogPath}:`, handleError);
                        errorCount++; item.modified = true; // Keep modified true on error
                        continue;
                    }
                }
                const writable = await tagHandle.createWritable();
                await writable.write(newTagString);
                await writable.close();
                item.originalTags = newTagString; // Update original tags baseline
                item.modified = false;
                savedCount++;
                document.getElementById(item.id)?.classList.remove('modified');
            } catch (writeError) {
                console.error(`Error saving file ${fullLogPath}:`, writeError);
                errorCount++; item.modified = true; // Keep modified true on error
            }
        }
        let finalMessage = `Saved ${savedCount} tag file(s).`;
        if (errorCount > 0) finalMessage += ` Failed to save ${errorCount}. Check console.`;
        
        updateSaveAllButtonState(); // Will disable if all modified items were saved
        recalculateAllTagStats(); // Stats might change if tags were consolidated etc.
        updateStatus(finalMessage, false);
    }

    function navigateEditor(direction) { // Operates on displayedImageDataIndices
        if (currentEditIndex === -1 || displayedImageDataIndices.length <= 1) return;

        const currentGlobalIndex = currentEditIndex; // This is an index in allImageData
        let currentDisplayIndexOfEditedItem = -1;
        for(let i=0; i < displayedImageDataIndices.length; i++){
            if(displayedImageDataIndices[i] === currentGlobalIndex){
                currentDisplayIndexOfEditedItem = i;
                break;
            }
        }
        
        if(currentDisplayIndexOfEditedItem === -1) {
             // Current edited item is not in the displayed list (e.g., filtered out after opening)
             // Try to find the closest visible item or just close editor.
             console.warn("Edited item not in current display for navigation. Closing editor.");
             closeEditor();
             return;
        }

        let nextItemDisplayIndex;
        if (direction === 'next') {
            nextItemDisplayIndex = (currentDisplayIndexOfEditedItem + 1) % displayedImageDataIndices.length;
        } else { // previous
            nextItemDisplayIndex = (currentDisplayIndexOfEditedItem - 1 + displayedImageDataIndices.length) % displayedImageDataIndices.length;
        }
        
        // Save current editor changes if any, before navigating
        const itemBeingEdited = allImageData[currentGlobalIndex];
        if (!itemBeingEdited.isTrashed && editorTagsTextarea.value !== currentEditOriginalTagsString) {
            saveEditorChanges(); // This also updates currentEditOriginalTagsString if successful
        }

        const nextGlobalIndexToOpen = displayedImageDataIndices[nextItemDisplayIndex];
        if (allImageData[nextGlobalIndexToOpen]) {
            openEditor(allImageData[nextGlobalIndexToOpen].id);
        }
    }

    function updateEditorNavButtonStates() {
        // Nav buttons are disabled if only one item is displayed OR if the current item is trashed (no nav from trashed)
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

        if (currentTagFragment.length === 0 && !value.endsWith(',')) { // Only hide if no fragment AND not ending with comma
            hideSuggestions();
            return;
        }

        const existingTagsInInput = parseTags(value.substring(0, lastCommaIndex + 1)); // Tags before current fragment
        
        // Suggestions from uniqueTags (which are based on active items)
        currentSuggestions = [...uniqueTags] 
            .filter(tag => 
                tag.toLowerCase().includes(currentTagFragment.toLowerCase()) && 
                !existingTagsInInput.includes(tag)
            )
            .sort((a,b) => a.localeCompare(b)) // Alphabetical sort
            .slice(0, 10); // Limit to 10 suggestions

        if (currentSuggestions.length > 0) {
            editorTagSuggestionsBox.innerHTML = ''; // Clear previous
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
            activeSuggestionIndex = -1; // Reset active suggestion
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
        // Move cursor to end
        suggestionInputTarget.selectionStart = suggestionInputTarget.selectionEnd = suggestionInputTarget.value.length;
    }

    function updateActiveSuggestion(newIndex) { // newIndex is for currentSuggestions array
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
        // Ensure all selected items are active (not trashed)
        const activeSelectedItemsData = [];
        selectedItemIds.forEach(id => {
            const item = getActiveItemById(id); // Use helper that checks isTrashed
            if (item) activeSelectedItemsData.push(item);
        });

        if (activeSelectedItemsData.length !== selectedItemIds.size) {
            updateStatus("Some selected items are trashed and cannot be batch edited. Deselect them or restore them.", false);
            // Optionally, auto-deselect trashed items here and proceed if any active ones remain
            // For now, just block if discrepancy.
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
        const allTagsInSelection = new Map(); // tag -> count

        activeSelectedItemsData.forEach((item, idx) => {
            const currentItemTags = new Set(item.tags);
            if (idx > 0) { // For common tags, intersect with previous
                commonTags = new Set([...commonTags].filter(tag => currentItemTags.has(tag)));
            }
            item.tags.forEach(tag => { // For all tags in selection
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
            if (!commonTags.has(tag)) { // Only show tags not in common
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

        selectedItemIds.forEach(id => { // Iterate over original selection IDs
            const index = findIndexById(id);
            if (index === -1) return; 
            const item = allImageData[index];
            if (item.isTrashed) return; // Skip trashed items, though openBatchEditor should prevent this state

            const originalItemTagsString = joinTags(item.tags);
            let currentItemTagsSet = new Set(item.tags);

            tagsToAdd.forEach(tag => currentItemTagsSet.add(tag));
            tagsToRemove.forEach(tag => currentItemTagsSet.delete(tag));
            
            const newItemTagsArray = [...currentItemTagsSet].sort(); // Convert set to sorted array
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
            if (overallStatsChanged) recalculateAllTagStats(); // Global recalculation
            batchEditorStatus.textContent = `Tag changes applied to ${modifiedCount} item(s). Remember to "Save All Changes".`;
            batchEditorStatus.style.color = 'var(--success-color)';
            batchEditorAddTagsInput.value = ''; // Clear inputs after successful application
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
            // Create a copy of selectedItemIds because moveToTrash modifies it
            const idsToTrash = new Set(selectedItemIds); 
            idsToTrash.forEach(itemId => {
                const item = allImageData.find(i => i.id === itemId);
                if (item && !item.isTrashed) {
                    moveToTrash(itemId); // This function handles individual UI and state updates
                    movedCount++;
                }
            });

            if (movedCount > 0) {
                updateStatus(`Moved ${movedCount} item(s) to trash.`);
            } else {
                updateStatus("No active items were moved to trash (they might have been already trashed).");
            }
            closeBatchEditor();
            filterAndSearch(); // Refresh grid view as items are now trashed
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
        if (customRules.length === 0) { // Add defaults if none exist
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
    function performBulkCustomRulesOperation() { // Operates on active items
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

        activeItems.forEach(item => { // Iterate only over active items
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
                case 'tag-frequency-search': event.preventDefault(); return; // Don't trigger anything
                case 'sort-property': case 'sort-order': buttonToClick = applySortBtn; break;
            }
            if (buttonToClick && !buttonToClick.disabled) { event.preventDefault(); buttonToClick.click(); return; }
        }
        
        const currentEditorItem = (currentEditIndex !== -1) ? allImageData[currentEditIndex] : null;
        if (isSingleEditorOpen && !isTextareaFocused && !isImageZoomed && (!currentEditorItem || !currentEditorItem.isTrashed)) { // Nav only if editor open AND not on textarea AND image not zoomed AND item not trashed
            if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
                event.preventDefault();
                navigateEditor(event.key === 'ArrowLeft' ? 'previous' : 'next');
                return;
            }
        }

        if (!isModalOpen && !isInputFocusedGeneral && displayedImageDataIndices.length > 0) {
            let newFocusDisplayIndex = keyboardFocusIndex; // keyboardFocusIndex is index in displayedImageDataIndices
            let handled = false;
            if (newFocusDisplayIndex === -1 && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(event.key)) {
                newFocusDisplayIndex = 0; handled = true;
            } else if (newFocusDisplayIndex !== -1) { // Only if already focused
                let columns = 1;
                const gridWidth = imageGrid.offsetWidth;
                const firstItemElement = imageGrid.querySelector('.grid-item:not(.trashed-item)'); // Use an active item for measurement
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
                        if (itemToOpen && !itemToOpen.isTrashed) { // Can only open non-trashed
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
        editorModalContent.classList.toggle('image-zoomed'); // Used to hide other content
    }

    function handleDocumentClickToClearSelection(event) {
        const isModalOpen = !editorModal.classList.contains('hidden') || 
                            !batchEditorModal.classList.contains('hidden') || 
                            !ruleEditorModal.classList.contains('hidden');
        if (isModalOpen) return;

        const clickedOnGridItem = event.target.closest('.grid-item');
        const clickedOnBatchEditBtn = event.target.closest('#batch-edit-selection-btn');
        // Don't clear selection if clicking on these specific elements
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
        filterAndSearch(); // Re-filter implies re-sort and re-render
    });

    renameTagBtn.addEventListener('click', () => performStandardBulkOperation('rename'));
    removeTagBtn.addEventListener('click', () => performStandardBulkOperation('remove'));
    addTagsBtn.addEventListener('click', () => performStandardBulkOperation('add'));
    removeDuplicatesBtn.addEventListener('click', () => performStandardBulkOperation('removeDuplicates'));
    
    findDuplicateFilesBtn.addEventListener('click', findDuplicateTagFiles);
    findOrphanedMissingBtn.addEventListener('click', findOrphanedAndMissingTagFiles);
    findCaseInconsistentTagsBtn.addEventListener('click', findCaseInconsistentTags);
    exportUniqueTagsBtn.addEventListener('click', exportUniqueTagList);

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
    loadCustomRules(); // Also populates dropdown
    sortPropertySelect.value = currentSortProperty;
    sortOrderSelect.value = currentSortOrder;
    updateStatus('Ready. Load a dataset folder.');
    updateSaveAllButtonState();
    updateEmptyTrashButtonState();

    if (!window.showDirectoryPicker) {
        updateStatus("Warning: Browser lacks File System Access API. Core functionality limited.");
        [loadDatasetBtn, saveAllBtn, emptyTrashBtn, batchEditSelectionBtn, applySortBtn, createNewRuleBtn, bulkApplyCustomRulesBtn,
         renameTagBtn, removeTagBtn, addTagsBtn, removeDuplicatesBtn, 
         findDuplicateFilesBtn, findOrphanedMissingBtn, findCaseInconsistentTagsBtn, exportUniqueTagsBtn]
        .forEach(btn => { if (btn) btn.disabled = true; });
    }
});