// ==UserScript==
// @name         D-DART (Driver Detention Analysis & Review Tool) v10.8
// @namespace    http://tampermonkey.net/
// @version      10.8
// @description  Enterprise Batch Driver Detention Analysis & Review Tool - Lane Validated
// @author       Sachin Vallakati
// @match        *://trans-logistics.amazon.com/*
// @match        *://smc-na-iad.iad.proxy.amazon.com/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_setClipboard
// @connect      smc-na-iad.iad.proxy.amazon.com
// @connect      smc.amazon.com
// @connect      us-east-1.prod.api.execution-tools.freight.amazon.dev
// @connect      trans-logistics.amazon.com
// @connect      amazon.sharepoint.com
// @connect      *
// @run-at       document-end
// ==/UserScript==

/**
 * @fileoverview D-DART - Driver Detention Analysis & Review Tool v10.8
 * Enterprise-grade tool for analyzing and processing driver detention charges
 * with dynamic SOW configuration from SharePoint and Settings Panel.
 *
 * @author Sachin Vallakati
 * @version 10.8
 * @license MIT
 *
 * @changelog v10.8
 * - Feature: Lane validation — only charges contracted lane (SMC commodities vs FMC stops)
 *   If SMC originStopLocationCode/destinationStopLocationCode ≠ FMC VR stop codes,
 *   releases holds with timestamp comment + lane mismatch comment. No charge.
 * - Fix #1: TokenManager.ensure() race condition (removed fetchLock, gate on fetchPromise)
 * - Fix #2: SMC write APIs (updateOrder, addPricingLine, addComment) now use
 *   HttpClient with retry/circuit-breaker via _authenticatedRequest helper.
 *   GMRequest.fetch now handles HTTP 409 as VERSION_CONFLICT.
 * - Fix #3: Settings escape listener no longer accumulates on repeated open/close
 * - Fix #4: Approval flow (handleApproved/handleDeclined) uses analysis.holdCode
 *   consistently instead of fragile string heuristic
 * - Fix #6: Cache cleanup interval stops when cache is empty (CPU savings)
 * - Fix #7: Keyboard shortcut listener stored and properly removed on destroy()
 * - Fix #9: Shipper expand/collapse toggles single card DOM (no full list re-render)
 * - Fix #11: Batch progress DOM refs cached (reduced getElementById calls per tick)
 * - Carries forward all v10.7 fixes (#1-48)
 */

(function() {
    'use strict';

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 1: TYPE DEFINITIONS (JSDoc)
     * ═══════════════════════════════════════════════════════════════════════════ */

    /**
     * @typedef {Object} AnalysisResult
     * @property {string} type - Result type from ResultType enum
     * @property {number} charge - Calculated charge amount
     * @property {string} breakdown - Breakdown description
     * @property {boolean} hitMax - Whether max charge was hit
     * @property {string} action - Action type from ActionType enum
     * @property {string} actionText - Display text for action
     * @property {string} comment - Comment to add to order
     * @property {boolean} hasHold - Whether hold exists
     * @property {string|null} holdCode - Hold pricing code
     * @property {boolean} detentionExists - Whether detention line exists
     * @property {number} existingAmount - Existing charge amount
     * @property {boolean} isPickup - Whether this is a pickup stop
     * @property {Object|null} fmcTimestamps - FMC timestamp data
     * @property {boolean} requiresApproval - Whether approval is required
     * @property {boolean} autoChargeAllowed - Whether auto-charge is allowed
     * @property {boolean} authNumberRequired - Whether auth number is required
     * @property {Object|null} sowConfig - SOW configuration
     * @property {boolean} processed - Whether action was processed
     * @property {string|null} processedAction - Type of processed action
     * @property {number|null} processedAmount - Processed amount
     * @property {string|null} processError - Error message if processing failed
     * @property {boolean} hasBorrowedTimestamps - Whether timestamps were borrowed
     * @property {string|null} approvalNote - Note about approval requirement
     * @property {number|null} calculatedAmount - Correct calculated amount (mismatch)
     * @property {number|null} mismatchAmount - Difference amount (mismatch)
     * @property {string|null} mismatchType - UNDER, OVER, CORRECT, SHOULD_NOT_EXIST
     * @property {string|null} orderStatusGroup - active, invoiced, paid
     * @property {boolean} laneMismatch - Whether lane mismatch was detected
     * @property {Object|null} laneValidation - Lane validation result
     */

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 2: CONFIGURATION
     * ═══════════════════════════════════════════════════════════════════════════ */

    const CONFIG = Object.freeze({
        VERSION: '10.8',
        APP_NAME: 'D-DART',
        APP_SUBTITLE: 'Production',
        AUTHOR: 'Sachin Vallakati (vallsach@)',

        SHAREPOINT: Object.freeze({
            SITE_URL: 'https://amazon.sharepoint.com/sites/DDART-Config',
            LIST_NAME: 'SOWConfiguration',
            API_URL: 'https://amazon.sharepoint.com/sites/DDART-Config/_api/web/lists/getbytitle(\'SOWConfiguration\')/items',
            REQUEST_TIMEOUT: 30000
        }),

        CACHE: Object.freeze({
            DURATION: 5 * 60 * 1000,
            MAX_SIZE: 200,
            SOW_DURATION: 24 * 60 * 60 * 1000,
            CLEANUP_INTERVAL: 60000
        }),

        UI: Object.freeze({
            TOAST_DURATION: 3000,
            PANEL_WIDTH: 520,
            PANEL_MIN_WIDTH: 56,
            COPY_POPUP_DURATION: 2000,
            SCROLL_DEBOUNCE: 100
        }),

        MODE: Object.freeze({
            STORAGE_KEY: 'ddart_operation_mode',
            ACTION: 'ACTION',
            REVIEW: 'REVIEW',
            DEFAULT: 'REVIEW'
        }),

        TOKEN: Object.freeze({
            MAX_AGE: 2 * 60 * 1000,
            WARNING_THRESHOLD: 30,
            CRITICAL_THRESHOLD: 10,
            STORAGE_KEY: 'ddart_csrf',
            TIME_KEY: 'ddart_csrf_time',
            REFRESH_INTERVAL: 60000,
            FETCH_TIMEOUT: 15000
        }),

        API: Object.freeze({
            REQUEST_TIMEOUT: 30000,
            MAX_RETRIES: 3,
            RETRY_BASE_DELAY: 300,
            RETRY_MAX_DELAY: 5000,
            RATE_LIMIT_MULTIPLIER: 2
        }),

        CIRCUIT_BREAKER: Object.freeze({
            FAILURE_THRESHOLD: 5,
            SUCCESS_THRESHOLD: 2,
            RESET_TIMEOUT: 30000,
            HALF_OPEN_TIMEOUT: 5000
        }),

        BATCH: Object.freeze({
            PARALLEL_SIZE: 5,
            MAX_ORDERS_PER_SESSION: 2000,
            CHUNK_SIZE: 50,
            CHUNK_DELAY: 1500,
            RATE_LIMIT_DELAY: 800,
            PROGRESS_SAVE_INTERVAL: 10,
            STORAGE_KEY: 'ddart_batch_progress',
            UI_UPDATE_INTERVAL: 300,
            PAUSE_CHECK_INTERVAL: 500,
            COOLDOWN_STATUS_DURATION: 1000
        }),

        APPROVAL: Object.freeze({
            TIMEOUT: 30000,
            WARNING_TIME: 10,
            CRITICAL_TIME: 5,
            COUNTDOWN_INTERVAL: 1000
        }),

        TIMING: Object.freeze({
            EARLY_MINUTES: -5,
            ON_TIME_MINUTES: 15,
            LATE_MINUTES: 0
        }),

        VALIDATION: Object.freeze({
            ORDER_ID_PATTERN: /^[A-Za-z0-9-_]+$/,
            ORDER_ID_MIN_LENGTH: 5,
            ORDER_ID_MAX_LENGTH: 50,
            AUTH_NUMBER_MAX_LENGTH: 100,
            MAX_CHARGE_AMOUNT: 10000
        }),

        URLS: Object.freeze({
            SMC_BASE: 'https://smc-na-iad.iad.proxy.amazon.com',
            SMC_ORDER: 'https://smc-na-iad.iad.proxy.amazon.com/order',
            SMC_EXECUTION_API: 'https://us-east-1.prod.api.execution-tools.freight.amazon.dev',
            FMC_BASE: 'https://trans-logistics.amazon.com',
            FMC_SEARCH: 'https://trans-logistics.amazon.com/fmc/execution/search'
        }),

        ALLOWED_DOMAINS: Object.freeze([
            'smc-na-iad.iad.proxy.amazon.com',
            'trans-logistics.amazon.com',
            'smc.amazon.com',
            'amazon.sharepoint.com'
        ]),

        KEYBOARD: Object.freeze({
            TOGGLE_MINIMIZE: 'Escape',
            ANALYZE: 'Enter',
            RESET: 'KeyR',
            SETTINGS: 'KeyS',
            DEBUG: 'KeyD'
        }),

        INITIAL_POSITION: Object.freeze({ top: '20px', left: '20px', right: 'auto' }),

        FEATURES: Object.freeze({
            TELEMETRY_ENABLED: true,
            KEYBOARD_SHORTCUTS: true
        }),

        START_MINIMIZED: true,
        MAX_DEBUG_LOGS: 300,
        DEBUG_ENABLED: true,

        PROGRESS: Object.freeze({ MAX_AGE: 7200000, SAVE_THROTTLE: 5000 }),
        TELEMETRY: Object.freeze({ MAX_EVENTS: 1000, FLUSH_INTERVAL: 60000 })
    });

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 3: CENTRALIZED MESSAGES
     * ═══════════════════════════════════════════════════════════════════════════ */

    const Messages = Object.freeze({
        ERRORS: Object.freeze({
            NO_DATA: 'No data to download',
            TOKEN_MISSING: 'Could not obtain CSRF token. Please log into SMC.',
            TOKEN_EXPIRED: 'CSRF token expired',
            INVALID_ORDER_IDS: 'Enter valid Order ID(s)',
            NETWORK_ERROR: 'Network connection failed. Please check your connection.',
            AUTH_ERROR: 'Authentication expired. Please refresh and log in again.',
            TIMEOUT_ERROR: 'Request timed out. Please try again.',
            PARSE_ERROR: 'Failed to process server response.',
            UNKNOWN_ERROR: 'An unexpected error occurred. Please try again.',
            VERSION_CONFLICT: 'Version conflict - order was modified',
            RESOURCE_NOT_FOUND: 'Resource not found',
            SOW_NOT_CONFIGURED: (shipper) => `SOW not configured for: ${shipper}`,
            SOW_DISABLED: (shipper) => `SOW is disabled for ${shipper}. Contact tool admin.`,
            SOW_INCOMPLETE: (shipper) => `Incomplete SOW configuration for: ${shipper}`,
            SOW_SERVER_UNREACHABLE: 'SOW Server Unreachable - Click to retry',
            SOW_AUTH_REQUIRED: 'Please login to SharePoint, then click Retry',
            AUTH_NUMBER_REQUIRED: 'Authorization number is required',
            COPY_FAILED: 'Failed to copy to clipboard',
            NO_EXECUTION_LEGS: 'No execution legs found in order',
            NO_TOUR_ID: 'No Tour ID found in execution leg',
            FMC_API_FAILURE: 'FMC API returned failure',
            EMPTY_RESPONSE: 'Empty response from API',
            BATCH_TOO_LARGE: (max) => `Maximum ${max} orders per session. Please split your batch.`,
            RATE_LIMITED: 'Rate limited by server, slowing down...',
            CIRCUIT_BREAKER_OPEN: 'Service temporarily unavailable. Please try again later.',
            HOLD_CODE_NOT_FOUND: (holdCode, orderId) => `Hold code ${holdCode} not found in pricing for ${orderId}`,
            HOLD_CODE_MISSING: 'No hold code available for action',
            LANE_MISMATCH: 'Lane mismatch — contracted lane does not match FMC execution'
        }),
        SUCCESS: Object.freeze({
            CSV_DOWNLOADED: 'CSV downloaded!',
            TXT_DOWNLOADED: 'TXT downloaded!',
            DEBUG_COPIED: 'Debug Log Copied!',
            TOKEN_FETCHED: 'Token fetched successfully',
            ORDER_UPDATED: 'Order updated successfully',
            CHARGE_ADDED: 'Charge added successfully',
            HOLD_RELEASED: 'Hold released successfully',
            SOW_LOADED: (count) => `SOW loaded: ${count} shipper(s)`,
            SOW_REFRESHED: 'SOW configuration refreshed',
            BATCH_COMPLETE: (success, failed) => `Batch complete: ${success} processed, ${failed} failed`,
            SETTINGS_SAVED: 'Settings saved',
            MODE_CHANGED: (mode) => `Switched to ${mode} Mode`
        }),
        INFO: Object.freeze({
            PROCESSING: 'Processing...',
            AWAITING_ARRIVAL: 'Awaiting arrival',
            AWAITING_DEPARTURE: 'Awaiting departure',
            NO_ACTION_NEEDED: 'No Action Needed',
            DRIVER_LATE: 'Driver late - No charge',
            DROP_HOOK: 'Drop & Hook - No detention',
            WITHIN_FREE_TIME: 'Within free time',
            BELOW_MINIMUM: 'Below minimum threshold',
            ANALYSIS_ONLY: 'Analysis only - Auto-charge disabled',
            BATCH_PAUSED: 'Batch paused - Click Resume to continue',
            BATCH_CANCELLED: 'Batch cancelled by user',
            TOKEN_REFRESHING: 'Refreshing authentication token...',
            COOLING_DOWN: 'Cooling down before next chunk...',
            SOW_LOADING: 'Loading SOW configuration...',
            REVIEW_MODE_NO_ACTIONS: 'Review Mode — No actions taken',
            REVIEW_MODE_BANNER: '📋 REVIEW MODE — No actions were taken. Results show what would happen in Action Mode.',
            APPROVAL_REQUIRED_NOTE: 'Approval required for this shipper',
            BORROWED_TIMESTAMP_NOTE: 'Timestamp issue (borrowed from tour)',
            SESSION_WARNING: 'Session may have expired. If errors occur, please refresh the page.',
            LANE_MISMATCH_DETECTED: 'Lane mismatch — Hold released'
        }),
        COMMENTS: Object.freeze({
            ADD_CHARGE: 'Driver Detention Charge Added',
            RELEASE_HOLD: 'As per FMC time stamps there is no delay for this load, there are no emails for the delay hence releasing the $0 DD charge.',
            CHARGE_WITH_AUTH: (authNumber) => `Driver Detention Charge Added - (${authNumber})`,
            APPROVAL_DECLINED: 'Shipper rejected DD charge, Releasing DD hold.',
            BORROWED_TIMESTAMP_RELEASE: 'Time stamps issue: After thorough investigation using all available tools and checking the order from tour level, we could not find any relevant information. Due to lack of required details, proceeding with releasing the hold',
            LANE_MISMATCH: (smcOrigin, smcDest, fmcOrigin, fmcDest) =>
                `Lane mismatch detected: SMC contracted lane (${smcOrigin} → ${smcDest}) does not match FMC execution lane (${fmcOrigin} → ${fmcDest}). Detention charge not applicable for this VR.`
        }),
        ACCESSIBILITY: Object.freeze({
            PANEL_LABEL: 'D-DART Driver Detention Analysis Tool',
            EXPAND_PANEL: 'Expand D-DART panel',
            MINIMIZE_PANEL: 'Minimize panel',
            OPEN_SETTINGS: 'Open settings',
            RESET_FORM: 'Reset form',
            COPY_DEBUG: 'Copy debug log',
            ANALYZE_ORDERS: 'Analyze orders',
            ORDER_INPUT: 'Order IDs input',
            RESULTS_REGION: 'Analysis results',
            CLOSE_SETTINGS: 'Close settings panel',
            APPROVE_CHARGE: 'Approve charge',
            DECLINE_CHARGE: 'Decline charge',
            SKIP_ORDER: 'Skip this order',
            TOGGLE_MODE: 'Toggle between Action and Review mode'
        })
    });

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 4: ENUMERATIONS
     * ═══════════════════════════════════════════════════════════════════════════ */

    const ResultType = Object.freeze({
        ORDER_CANCELLED: 'ORDER_CANCELLED',
        ORDER_INVOICED: 'ORDER_INVOICED',
        MISSING_ARRIVAL: 'MISSING_ARRIVAL',
        MISSING_DEPARTURE: 'MISSING_DEPARTURE',
        NO_DETENTION_DROP_HOOK: 'NO_DETENTION_DROP_HOOK',
        DRIVER_LATE: 'DRIVER_LATE',
        WITHIN_FREE_TIME: 'WITHIN_FREE_TIME',
        BELOW_MINIMUM_THRESHOLD: 'BELOW_MINIMUM_THRESHOLD',
        CHARGEABLE: 'CHARGEABLE',
        CHARGEABLE_MAX: 'CHARGEABLE_MAX',
        CHARGE_EXISTS: 'CHARGE_EXISTS',
        NO_HOLD_NO_CHARGE: 'NO_HOLD_NO_CHARGE',
        FMC_DATA_UNAVAILABLE: 'FMC_DATA_UNAVAILABLE',
        SOW_NOT_CONFIGURED: 'SOW_NOT_CONFIGURED',
        SOW_DISABLED: 'SOW_DISABLED',
        SOW_INCOMPLETE: 'SOW_INCOMPLETE',
        UNKNOWN_ERROR: 'UNKNOWN_ERROR',
        CHARGE_MISMATCH_UNDER: 'CHARGE_MISMATCH_UNDER',
        CHARGE_MISMATCH_OVER: 'CHARGE_MISMATCH_OVER',
        CHARGE_CORRECT: 'CHARGE_CORRECT',
        CHARGE_SHOULD_NOT_EXIST: 'CHARGE_SHOULD_NOT_EXIST',
        LANE_MISMATCH: 'LANE_MISMATCH'
    });

    const ActionType = Object.freeze({
        ADD_CHARGE_UPDATE: 'ADD_CHARGE_UPDATE',
        ADD_CHARGE_CREATE: 'ADD_CHARGE_CREATE',
        RELEASE_HOLD: 'RELEASE_HOLD',
        CREATE_HOLD_ONLY: 'CREATE_HOLD_ONLY',
        ANALYSIS_ONLY: 'ANALYSIS_ONLY',
        NO_ACTION: 'NO_ACTION',
        PENDING: 'PENDING',
        ERROR: 'ERROR',
        PENDING_APPROVAL: 'PENDING_APPROVAL',
        SKIPPED: 'SKIPPED'
    });

    const ErrorType = Object.freeze({
        NETWORK: 'NETWORK_ERROR',
        AUTH: 'AUTH_ERROR',
        VALIDATION: 'VALIDATION_ERROR',
        BUSINESS: 'BUSINESS_ERROR',
        TIMEOUT: 'TIMEOUT_ERROR',
        PARSE: 'PARSE_ERROR',
        RATE_LIMIT: 'RATE_LIMIT_ERROR',
        CIRCUIT_BREAKER: 'CIRCUIT_BREAKER_ERROR',
        SOW: 'SOW_ERROR',
        UNKNOWN: 'UNKNOWN_ERROR'
    });

    const BatchState = Object.freeze({
        IDLE: 'IDLE', RUNNING: 'RUNNING', PAUSED: 'PAUSED',
        CANCELLED: 'CANCELLED', COMPLETED: 'COMPLETED'
    });

    const CircuitBreakerState = Object.freeze({
        CLOSED: 'CLOSED', OPEN: 'OPEN', HALF_OPEN: 'HALF_OPEN'
    });

    const SOWStatus = Object.freeze({
        NOT_LOADED: 'NOT_LOADED', LOADING: 'LOADING', LOADED: 'LOADED',
        ERROR: 'ERROR', AUTH_REQUIRED: 'AUTH_REQUIRED'
    });

    const ShipperStatus = Object.freeze({
        ACTIVE: 'ACTIVE', INACTIVE: 'INACTIVE',
        VALIDATION_ERROR: 'VALIDATION_ERROR', INCOMPLETE: 'INCOMPLETE'
    });

    const TelemetryEventType = Object.freeze({
        APP_INIT: 'APP_INIT', APP_ERROR: 'APP_ERROR',
        BATCH_START: 'BATCH_START', BATCH_COMPLETE: 'BATCH_COMPLETE',
        ORDER_PROCESSED: 'ORDER_PROCESSED', TOKEN_REFRESH: 'TOKEN_REFRESH',
        SOW_LOAD: 'SOW_LOAD', USER_ACTION: 'USER_ACTION'
    });

    const OrderStatusMap = Object.freeze({
        'IN_DRAFT': { display: 'Draft', color: '#337ab7', group: 'draft' },
        'NOT_PLANNED': { display: 'Uncovered', color: '#337ab7', group: 'planning' },
        'PENDING_CARRIER_ACCEPTANCE': { display: 'Tendered', color: '#f0ad4e', group: 'active' },
        'CARRIER_TENDER_ACCEPTED': { display: 'Covered', color: '#f0ad4e', group: 'active' },
        'DRIVER_DISPATCHED': { display: 'Dispatched', color: '#f0ad4e', group: 'active' },
        'LATE_TO_ARRIVE': { display: 'Late to Arrive', color: '#f0ad4e', group: 'active' },
        'ARRIVED': { display: 'Arrived', color: '#f0ad4e', group: 'active' },
        'LATE_TO_DEPART': { display: 'Late to Depart', color: '#f0ad4e', group: 'active' },
        'DEPARTED': { display: 'Departed', color: '#f0ad4e', group: 'active' },
        'PENDING_DELIVERY_CONFIRMATION': { display: 'Pending POD', color: '#f0ad4e', group: 'active' },
        'DELIVERY_CONFIRMED': { display: 'Pending Invoicing', color: '#f0ad4e', group: 'complete' },
        'PENDING_PAYMENT': { display: 'Invoiced', color: '#5cb85c', group: 'invoiced' },
        'PAID': { display: 'Paid', color: '#5cb85c', group: 'paid' },
        'CANCELLED': { display: 'Cancelled', color: '#d9534f', group: 'cancelled' },
        'REJECTED': { display: 'Rejected', color: '#d9534f', group: 'cancelled' }
    });

    const DetentionPricing = Object.freeze({
        SHIPPER: Object.freeze({ pricingCode: 'DETENTION_DRIVER_AT_SHIPPER', description: 'Driver Detention at Shipper Charge' }),
        RECEIVER: Object.freeze({ pricingCode: 'DETENTION_DRIVER_AT_RECEIVER', description: 'Driver Detention at Receiver Charge' })
    });

    const ActionDisplayConfig = Object.freeze({
        CHARGE_ADDED: Object.freeze({ term: 'Charge Added', icon: '✅', display: (amount) => `✅ Charge Added - $${amount.toFixed(2)}`, reportTerm: 'Charge Added', cssClass: 'added' }),
        RECOVERED: Object.freeze({ term: '🎯 Recovered', icon: '🎯', display: (amount) => `🎯 Recovered - $${amount.toFixed(2)}`, reportTerm: '🎯 Recovered', cssClass: 'recovered' }),
        HOLD_RELEASED: Object.freeze({ term: 'Hold Released', icon: '✅', display: () => '✅ Hold Released', reportTerm: 'Hold Released', cssClass: 'released' }),
        HOLD_CREATED: Object.freeze({ term: 'Hold Created', icon: '📋', display: () => '📋 $0 Hold Created (Analysis Only)', reportTerm: 'Hold Created', cssClass: 'hold-created' }),
        ANALYSIS_ONLY: Object.freeze({ term: 'Analysis Only', icon: '📊', display: (amount) => `📊 Analysis Only - $${amount ? amount.toFixed(2) : '0.00'}`, reportTerm: 'Analysis Only', cssClass: 'analysis' }),
        PENDING: Object.freeze({ term: 'Pending', icon: '⏳', display: (reason) => `⏳ Pending${reason ? ' - ' + reason : ''}`, reportTerm: 'Pending', cssClass: 'pending' }),
        NO_ACTION: Object.freeze({ term: 'No Action', icon: '—', display: () => '— No Action', reportTerm: 'No Action', cssClass: 'no-action' }),
        ERROR: Object.freeze({ term: 'Error', icon: '❌', display: (message) => `❌ Error${message ? ' - ' + message : ''}`, reportTerm: 'Error', cssClass: 'error' }),
        LANE_MISMATCH: Object.freeze({ term: 'Lane Mismatch', icon: '🚫', display: () => '🚫 Lane Mismatch — Hold Released', reportTerm: 'Lane Mismatch', cssClass: 'lane-mismatch' }),
        CHARGE_MISMATCH: Object.freeze({ term: 'Charge Mismatch', icon: '⚠️', display: (details) => `⚠️ ${details}`, reportTerm: 'Charge Mismatch', cssClass: 'mismatch' }),
        CHARGE_CORRECT: Object.freeze({ term: 'Charge Correct', icon: '✅', display: (amount) => `✅ Charge Correct ($${amount.toFixed(2)})`, reportTerm: 'Charge Correct', cssClass: 'correct' })
    });

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 5: UTILITY FUNCTIONS
     * ═══════════════════════════════════════════════════════════════════════════ */

    const generateSessionId = () => `${Date.now().toString(36)}-${Math.random().toString(36).substring(2, 11)}`;

    const safeStringify = (obj, indent = 2) => {
        const seen = new WeakSet();
        return JSON.stringify(obj, (key, value) => {
            if (value instanceof Map) return { __type: 'Map', entries: Array.from(value.entries()) };
            if (value instanceof Set) return { __type: 'Set', values: Array.from(value.values()) };
            if (typeof value === 'object' && value !== null) {
                if (seen.has(value)) return '[Circular Reference]';
                seen.add(value);
            }
            return value;
        }, indent);
    };

    const parseBoolean = (value, defaultValue = false) => {
        if (value === null || value === undefined || value === '') return defaultValue;
        if (typeof value === 'boolean') return value;
        if (typeof value === 'string') {
            const normalized = value.toLowerCase().trim();
            if (['true', 'yes', '1', 'on', 'enabled'].includes(normalized)) return true;
            if (['false', 'no', '0', 'off', 'disabled'].includes(normalized)) return false;
        }
        if (typeof value === 'number') return value !== 0;
        return defaultValue;
    };

    const debounce = (func, wait, immediate = false) => {
        let timeout;
        return function executedFunction(...args) {
            const context = this;
            const later = () => { timeout = null; if (!immediate) func.apply(context, args); };
            const callNow = immediate && !timeout;
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
            if (callNow) func.apply(context, args);
        };
    };

    const chunkArray = (array, size) => {
        const chunks = [];
        for (let i = 0; i < array.length; i += size) chunks.push(array.slice(i, i + size));
        return chunks;
    };

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 6: TELEMETRY SERVICE
     * ═══════════════════════════════════════════════════════════════════════════ */

    const Telemetry = (() => {
        const events = [];
        const sessionId = generateSessionId();
        let flushTimeout = null;
        const isEnabled = CONFIG.FEATURES.TELEMETRY_ENABLED;

        const track = (eventType, data = {}) => {
            if (!isEnabled) return;
            const event = { event: eventType, data: { ...data, version: CONFIG.VERSION, url: window.location.hostname }, timestamp: Date.now(), sessionId };
            events.push(event);
            while (events.length > CONFIG.TELEMETRY.MAX_EVENTS) events.shift();
            if (!flushTimeout) {
                flushTimeout = setTimeout(() => { flush(); flushTimeout = null; }, CONFIG.TELEMETRY.FLUSH_INTERVAL);
            }
        };

        const flush = () => {
            if (CONFIG.DEBUG_ENABLED && events.length > 0) {
                console.log(`🚛 D-DART Telemetry: ${events.length} events in session ${sessionId}`);
            }
        };

        const getEvents = () => [...events];
        const getSessionId = () => sessionId;

        const getMetrics = () => {
            const metrics = { sessionId, eventCount: events.length, sessionDuration: Date.now() - (events[0]?.timestamp || Date.now()), eventTypes: {} };
            events.forEach(e => { metrics.eventTypes[e.event] = (metrics.eventTypes[e.event] || 0) + 1; });
            return metrics;
        };

        const cleanup = () => {
            if (flushTimeout) { clearTimeout(flushTimeout); flushTimeout = null; }
            flush();
        };

        return { track, flush, getEvents, getSessionId, getMetrics, cleanup };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 7: PERFORMANCE MONITOR
     * ═══════════════════════════════════════════════════════════════════════════ */

    const PerformanceMonitor = (() => {
        const metrics = { apiCalls: 0, apiErrors: 0, avgResponseTime: 0, totalResponseTime: 0, ordersProcessed: 0, cacheHits: 0, cacheMisses: 0 };

        const recordApiCall = (duration, isError = false) => {
            metrics.apiCalls++;
            if (isError) metrics.apiErrors++;
            else { metrics.totalResponseTime += duration; metrics.avgResponseTime = metrics.totalResponseTime / (metrics.apiCalls - metrics.apiErrors); }
        };

        const recordOrderProcessed = () => { metrics.ordersProcessed++; };
        const recordCacheHit = () => { metrics.cacheHits++; };
        const recordCacheMiss = () => { metrics.cacheMisses++; };

        const getMetrics = () => ({
            ...metrics,
            cacheHitRate: metrics.cacheHits + metrics.cacheMisses > 0 ? (metrics.cacheHits / (metrics.cacheHits + metrics.cacheMisses) * 100).toFixed(2) + '%' : 'N/A',
            apiSuccessRate: metrics.apiCalls > 0 ? ((metrics.apiCalls - metrics.apiErrors) / metrics.apiCalls * 100).toFixed(2) + '%' : 'N/A'
        });

        return { recordApiCall, recordOrderProcessed, recordCacheHit, recordCacheMiss, getMetrics };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 8: STATE MANAGEMENT
     * ═══════════════════════════════════════════════════════════════════════════ */

    class StateManager {
        constructor() {
            this._state = {
                batchReportData: [], pendingApprovalOrders: [],
                isProcessing: false, currentOrderIds: [], currentIndex: 0, totalOrders: 0,
                batchState: BatchState.IDLE, processedOrders: new Map(), failedOrders: [],
                batchStartTime: null, currentChunk: 0, totalChunks: 0,
                isMinimized: CONFIG.START_MINIMIZED, isSingleMode: false, singleOrderData: null,
                operationMode: CONFIG.MODE.DEFAULT,
                sowStatus: SOWStatus.NOT_LOADED, sowShipperCount: 0, sowLastError: null, sowLastRefresh: null,
                isSettingsOpen: false, expandedShippers: new Set(),
                _previousResultsContent: ''
            };
            this._listeners = new Map();
        }

        get(key) { return this._state[key]; }

        set(key, value) {
            const oldValue = this._state[key];
            this._state[key] = value;
            this._notify(key, value, oldValue);
        }

        update(updates) {
            const entries = Object.entries(updates);
            for (let i = 0; i < entries.length; i++) { const [key, value] = entries[i]; this.set(key, value); }
        }

        subscribe(key, callback) {
            if (!this._listeners.has(key)) this._listeners.set(key, new Set());
            this._listeners.get(key).add(callback);
            return () => { const listeners = this._listeners.get(key); if (listeners) listeners.delete(callback); };
        }

        resetBatch() {
            this.update({
                batchReportData: [], pendingApprovalOrders: [],
                currentOrderIds: [], currentIndex: 0, totalOrders: 0,
                batchState: BatchState.IDLE, processedOrders: new Map(), failedOrders: [],
                batchStartTime: null, currentChunk: 0, totalChunks: 0,
                isSingleMode: false, singleOrderData: null
            });
        }

        addBatchReportEntry(entry) { const current = this._state.batchReportData; this.set('batchReportData', [...current, entry]); }
        addPendingApprovalOrder(orderData) { const current = this._state.pendingApprovalOrders; this.set('pendingApprovalOrders', [...current, orderData]); }
        addProcessedOrder(orderId, data) { const map = new Map(this._state.processedOrders); map.set(orderId, data); this.set('processedOrders', map); }
        addFailedOrder(orderId, error) { const current = this._state.failedOrders; this.set('failedOrders', [...current, { orderId, error, timestamp: Date.now() }]); }

        toggleShipperExpanded(shipperName) {
            const expanded = new Set(this._state.expandedShippers);
            if (expanded.has(shipperName)) expanded.delete(shipperName);
            else expanded.add(shipperName);
            this.set('expandedShippers', expanded);
        }

        _notify(key, newValue, oldValue) {
            const listeners = this._listeners.get(key);
            if (listeners) {
                listeners.forEach(callback => {
                    try { callback(newValue, oldValue); }
                    catch (error) { console.error('State listener error', { key, error: error.message }); }
                });
            }
        }

        getSnapshot() {
            return {
                ...this._state,
                processedOrders: Array.from(this._state.processedOrders.entries()),
                processedOrdersCount: this._state.processedOrders.size,
                failedOrdersCount: this._state.failedOrders.length,
                expandedShippers: Array.from(this._state.expandedShippers),
                expandedShippersCount: this._state.expandedShippers.size
            };
        }

        clearListeners() { this._listeners.clear(); }
    }

    const AppState = new StateManager();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 9: LOGGING SERVICE
     * ═══════════════════════════════════════════════════════════════════════════ */

    const Logger = (() => {
        const logs = [];
        const errorCounts = new Map();

        const addLog = (level, message, data = null) => {
            if (!CONFIG.DEBUG_ENABLED && level === 'DEBUG') return;
            const timestamp = new Date().toISOString();
            let logEntry = `[${timestamp}] [${level}] ${message}`;
            if (data !== null) {
                try {
                    const dataStr = typeof data === 'string' ? data : safeStringify(data);
                    const truncated = dataStr.length > 1000 ? dataStr.substring(0, 1000) + '...[truncated]' : dataStr;
                    logEntry += `\n  DATA: ${truncated}`;
                } catch (e) { logEntry += `\n  DATA: [Could not stringify: ${e.message}]`; }
            }
            logs.unshift(logEntry);
            if (logs.length > CONFIG.MAX_DEBUG_LOGS) logs.pop();
            if (level === 'ERROR') {
                const key = message.substring(0, 50);
                errorCounts.set(key, (errorCounts.get(key) || 0) + 1);
                Telemetry.track(TelemetryEventType.APP_ERROR, { message: message.substring(0, 200), level });
            }
            const consoleMethod = level === 'ERROR' ? 'error' : level === 'WARN' ? 'warn' : 'log';
            console[consoleMethod](`🚛 D-DART [${level}]:`, message, data || '');
        };

        return {
            debug: (message, data) => addLog('DEBUG', message, data),
            info: (message, data) => addLog('INFO', message, data),
            warn: (message, data) => addLog('WARN', message, data),
            error: (message, data) => addLog('ERROR', message, data),

            generateReport: () => {
                const performanceMetrics = PerformanceMonitor.getMetrics();
                const telemetryMetrics = Telemetry.getMetrics();
                return safeStringify({
                    '=== D-DART DEBUG REPORT ===': new Date().toISOString(),
                    'Version': CONFIG.VERSION, 'Edition': CONFIG.APP_SUBTITLE,
                    'Page URL': window.location.href, 'User Agent': navigator.userAgent,
                    'Is on SMC': Helpers.isOnSMC(), 'Operation Mode': ModeManager.getMode(),
                    'Token Available': !!TokenManager.getToken(), 'Token Age': TokenManager.getAgeString(),
                    'Token Remaining': TokenManager.getRemainingSeconds() + 's',
                    'SOW Status': AppState.get('sowStatus'), 'SOW Shipper Count': AppState.get('sowShipperCount'),
                    'State Snapshot': AppState.getSnapshot(),
                    'Circuit Breakers': { smc: circuitBreakers.smc.getState(), fmc: circuitBreakers.fmc.getState(), execution: circuitBreakers.execution.getState() },
                    'Performance Metrics': performanceMetrics, 'Telemetry Metrics': telemetryMetrics,
                    'Error Frequency': Object.fromEntries(errorCounts), 'Health Check': HealthCheck.check(),
                    'Recent Logs': logs.slice(0, 100)
                });
            },

            getLogs: () => [...logs]
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 10: UTILITY HELPERS
     * ═══════════════════════════════════════════════════════════════════════════ */

    const Helpers = {
        escapeHtml(text) {
            if (text == null) return '';
            const div = document.createElement('div');
            div.textContent = String(text);
            return div.innerHTML;
        },

        generateId(prefix = 'id') {
            return `${prefix}-${Math.random().toString(36).substring(2, 11)}-${Date.now().toString(36)}`;
        },

        isOnSMC() {
            const host = window.location.hostname;
            return host.includes('smc-na-iad') || host.includes('smc.amazon.com');
        },

        formatDuration(minutes) {
            if (typeof minutes !== 'number' || isNaN(minutes)) return 'N/A';
            const absMinutes = Math.abs(minutes);
            const hours = Math.floor(absMinutes / 60);
            const mins = absMinutes % 60;
            return hours > 0 ? `${hours}h ${mins}m` : `${mins}m`;
        },

        formatCurrency(amount) {
            if (typeof amount !== 'number' || isNaN(amount)) return '$0.00';
            return `$${amount.toFixed(2)}`;
        },

        formatTimeFromEpoch(epochMs, timezone = 'America/Los_Angeles') {
            if (!epochMs) return '-';
            try {
                return new Date(epochMs).toLocaleString('en-US', {
                    timeZone: timezone, month: 'short', day: 'numeric',
                    hour: '2-digit', minute: '2-digit', hour12: false
                });
            } catch (e) { return new Date(epochMs).toLocaleString(); }
        },

        formatDelta(deltaMinutes) {
            if (deltaMinutes === null || deltaMinutes === undefined) return '-';
            const absMinutes = Math.abs(deltaMinutes);
            const sign = deltaMinutes >= 0 ? '+' : '-';
            if (absMinutes >= 60) { const hours = Math.floor(absMinutes / 60); const mins = absMinutes % 60; return `${sign}${hours}h ${mins}m`; }
            return `${sign}${absMinutes}m`;
        },

        formatETA(ms) {
            if (ms < 60000) return '< 1 minute';
            else if (ms < 3600000) return `~${Math.ceil(ms / 60000)} minutes`;
            else { const hours = Math.floor(ms / 3600000); const mins = Math.ceil((ms % 3600000) / 60000); return `~${hours}h ${mins}m`; }
        },

        formatElapsed(ms) {
            if (ms < 60000) return `${Math.round(ms / 1000)}s`;
            else if (ms < 3600000) return `${Math.floor(ms / 60000)}m ${Math.round((ms % 60000) / 1000)}s`;
            else { const hours = Math.floor(ms / 3600000); const mins = Math.floor((ms % 3600000) / 60000); return `${hours}h ${mins}m`; }
        },

        formatRelativeTime(timestamp) {
            if (!timestamp) return 'Never';
            const diff = Date.now() - timestamp;
            if (diff < 60000) return 'Just now';
            if (diff < 3600000) return `${Math.floor(diff / 60000)} min ago`;
            if (diff < 86400000) return `${Math.floor(diff / 3600000)} hour(s) ago`;
            return `${Math.floor(diff / 86400000)} day(s) ago`;
        },

        truncateText(text, maxLength = 30) {
            if (!text || text.length <= maxLength) return text || '';
            return text.substring(0, maxLength - 3) + '...';
        },

        async copyToClipboard(text) {
            try { if (typeof GM_setClipboard === 'function') { GM_setClipboard(text); return; } } catch (e) { /* Fall through */ }
            if (navigator.clipboard?.writeText) return navigator.clipboard.writeText(text);
            return new Promise((resolve, reject) => {
                try {
                    const textarea = document.createElement('textarea');
                    textarea.value = text;
                    textarea.style.cssText = 'position:fixed;opacity:0;pointer-events:none;left:-9999px';
                    document.body.appendChild(textarea); textarea.select();
                    const success = document.execCommand('copy');
                    document.body.removeChild(textarea);
                    success ? resolve() : reject(new Error(Messages.ERRORS.COPY_FAILED));
                } catch (e) { reject(e); }
            });
        },

        formatBorrowedIndicator(sourceVrId) {
            if (!sourceVrId) return '';
            const shortVrId = sourceVrId.length > 15 ? sourceVrId.substring(0, 15) + '...' : sourceVrId;
            return `<div class="d-dart-borrowed-indicator">📍 ${this.escapeHtml(shortVrId)}</div>`;
        },

        downloadFile(content, filename, mimeType) {
            const blob = new Blob([content], { type: mimeType });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = filename;
            document.body.appendChild(a); a.click();
            document.body.removeChild(a); URL.revokeObjectURL(url);
        },

        formatValueOrDash(value) {
            if (value === null || value === undefined || value === '') return '-';
            return String(value);
        },

        formatBoolean(value) {
            if (parseBoolean(value, null) === true) return 'Yes';
            if (parseBoolean(value, null) === false) return 'No';
            return '-';
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 11: URL SECURITY
     * ═══════════════════════════════════════════════════════════════════════════ */

    const SecurityHelpers = {
        sanitizeUrl(url) {
            if (!url) return '#';
            try {
                const parsed = new URL(url);
                if (parsed.protocol !== 'https:') return '#';
                const isAllowed = CONFIG.ALLOWED_DOMAINS.some(domain => parsed.hostname.includes(domain));
                if (!isAllowed) return '#';
                return url;
            } catch (e) { return '#'; }
        },

        buildSMCOrderUrl(orderId) {
            if (!orderId) return '#';
            return this.sanitizeUrl(`${CONFIG.URLS.SMC_ORDER}/${encodeURIComponent(orderId)}`);
        },

        buildFMCSearchUrl(searchId) {
            if (!searchId) return '#';
            return this.sanitizeUrl(`${CONFIG.URLS.FMC_SEARCH}/${encodeURIComponent(searchId)}`);
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 12: VALIDATION SERVICE
     * ═══════════════════════════════════════════════════════════════════════════ */

    const Validator = {
        isValidOrderId(id) {
            if (!id || typeof id !== 'string') return false;
            const trimmed = id.trim();
            if (trimmed.length < CONFIG.VALIDATION.ORDER_ID_MIN_LENGTH || trimmed.length > CONFIG.VALIDATION.ORDER_ID_MAX_LENGTH) return false;
            return CONFIG.VALIDATION.ORDER_ID_PATTERN.test(trimmed);
        },

        parseOrderIds(input) {
            const result = { valid: true, errors: [], sanitized: [], duplicatesRemoved: 0 };
            if (!input || typeof input !== 'string') { result.valid = false; result.errors.push('Input is required'); return result; }
            const rawIds = input.split(/[,\s\n]+/).map(id => id.trim()).filter(id => id.length > 0);
            if (rawIds.length === 0) { result.valid = false; result.errors.push('No order IDs found in input'); return result; }
            const uniqueIds = [...new Set(rawIds)];
            result.duplicatesRemoved = rawIds.length - uniqueIds.length;
            for (let i = 0; i < uniqueIds.length; i++) {
                const id = uniqueIds[i];
                if (this.isValidOrderId(id)) result.sanitized.push(id);
                else result.errors.push(`Invalid order ID: ${id}`);
            }
            if (result.sanitized.length === 0) result.valid = false;
            return result;
        },

        isValidAuthNumber(authNumber) {
            if (!authNumber || typeof authNumber !== 'string') return false;
            const trimmed = authNumber.trim();
            return trimmed.length > 0 && trimmed.length <= CONFIG.VALIDATION.AUTH_NUMBER_MAX_LENGTH;
        },

        sanitizeAuthNumber(authNumber) {
            if (!authNumber || typeof authNumber !== 'string') return null;
            const trimmed = authNumber.trim();
            return trimmed.length > 0 ? trimmed.substring(0, CONFIG.VALIDATION.AUTH_NUMBER_MAX_LENGTH) : null;
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 13: CACHE MANAGER (Fix #6: stop interval when empty)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const CacheManager = (() => {
        const cache = new Map();
        let cleanupInterval = null;

        const startCleanup = () => {
            if (cleanupInterval) return;
            cleanupInterval = setInterval(() => { cleanup(); }, CONFIG.CACHE.CLEANUP_INTERVAL);
        };

        const stopCleanup = () => { if (cleanupInterval) { clearInterval(cleanupInterval); cleanupInterval = null; } };

        const cleanup = () => {
            const now = Date.now(); let cleaned = 0;
            for (const [key, value] of cache.entries()) {
                const duration = value.duration || CONFIG.CACHE.DURATION;
                if (now - value.time > duration) { cache.delete(key); cleaned++; }
            }
            if (cleaned > 0) Logger.debug(`Cache cleanup: removed ${cleaned} expired entries`);
            // Fix #6: stop interval when cache is empty
            if (cache.size === 0) stopCleanup();
        };

        return {
            add(id, data, customDuration = null) {
                if (cache.size >= CONFIG.CACHE.MAX_SIZE) {
                    let oldestKey = null, oldestTime = Infinity;
                    for (const [key, value] of cache.entries()) { if (value.time < oldestTime) { oldestTime = value.time; oldestKey = key; } }
                    if (oldestKey) cache.delete(oldestKey);
                }
                cache.set(id, { data, time: Date.now(), duration: customDuration || CONFIG.CACHE.DURATION });
                startCleanup();
            },

            get(id) {
                const cached = cache.get(id);
                if (cached) {
                    const duration = cached.duration || CONFIG.CACHE.DURATION;
                    if (Date.now() - cached.time < duration) { PerformanceMonitor.recordCacheHit(); return cached.data; }
                    cache.delete(id);
                }
                PerformanceMonitor.recordCacheMiss();
                return null;
            },

            invalidate(id) { cache.delete(id); },
            clear() { cache.clear(); stopCleanup(); },
            getStats() { return { size: cache.size, maxSize: CONFIG.CACHE.MAX_SIZE, keys: Array.from(cache.keys()) }; },
            cleanup() { stopCleanup(); cache.clear(); }
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 14: PROGRESS PERSISTENCE
     * ═══════════════════════════════════════════════════════════════════════════ */

    const ProgressManager = (() => {
        let saveThrottleTimeout = null;
        return {
            save(data) {
                if (saveThrottleTimeout) return true;
                try {
                    const payload = { ...data, timestamp: Date.now(), version: CONFIG.VERSION, sessionId: Telemetry.getSessionId() };
                    GM_setValue(CONFIG.BATCH.STORAGE_KEY, safeStringify(payload));
                    Logger.debug('Progress saved', { orders: data.processedCount });
                    saveThrottleTimeout = setTimeout(() => { saveThrottleTimeout = null; }, CONFIG.PROGRESS.SAVE_THROTTLE);
                    return true;
                } catch (e) { Logger.warn('Failed to save progress', e.message); return false; }
            },

            load() {
                try {
                    const saved = GM_getValue(CONFIG.BATCH.STORAGE_KEY, null);
                    if (!saved) return null;
                    const data = JSON.parse(saved);
                    if (Date.now() - data.timestamp > CONFIG.PROGRESS.MAX_AGE) { this.clear(); return null; }
                    if (data.version !== CONFIG.VERSION) { Logger.info('Progress from different version, clearing'); this.clear(); return null; }
                    return data;
                } catch (e) { Logger.warn('Failed to load progress', e.message); return null; }
            },

            clear() {
                try { GM_setValue(CONFIG.BATCH.STORAGE_KEY, null); if (saveThrottleTimeout) { clearTimeout(saveThrottleTimeout); saveThrottleTimeout = null; } }
                catch (e) { Logger.warn('Failed to clear progress', e.message); }
            }
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 14.5: MODE MANAGER
     * ═══════════════════════════════════════════════════════════════════════════ */

    const ModeManager = (() => {
        return {
            init() {
                try {
                    const saved = GM_getValue(CONFIG.MODE.STORAGE_KEY, CONFIG.MODE.DEFAULT);
                    const mode = (saved === CONFIG.MODE.ACTION || saved === CONFIG.MODE.REVIEW) ? saved : CONFIG.MODE.DEFAULT;
                    AppState.set('operationMode', mode);
                    Logger.info(`Mode initialized: ${mode}`);
                } catch (e) { AppState.set('operationMode', CONFIG.MODE.DEFAULT); Logger.warn('Failed to load mode, defaulting to REVIEW', e.message); }
            },

            toggle() {
                const current = AppState.get('operationMode');
                const newMode = current === CONFIG.MODE.ACTION ? CONFIG.MODE.REVIEW : CONFIG.MODE.ACTION;
                AppState.set('operationMode', newMode);
                try { GM_setValue(CONFIG.MODE.STORAGE_KEY, newMode); } catch (e) { Logger.warn('Failed to persist mode', e.message); }
                Logger.info(`Mode switched to: ${newMode}`);
                Telemetry.track(TelemetryEventType.USER_ACTION, { action: 'mode_toggle', mode: newMode });
                if (typeof UIController !== 'undefined' && UIController.showToast) UIController.showToast(Messages.SUCCESS.MODE_CHANGED(newMode), 'success');
            },

            isReviewMode() { return AppState.get('operationMode') === CONFIG.MODE.REVIEW; },
            isActionMode() { return AppState.get('operationMode') === CONFIG.MODE.ACTION; },
            getMode() { return AppState.get('operationMode'); }
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 15: CIRCUIT BREAKER PATTERN
     * ═══════════════════════════════════════════════════════════════════════════ */

    class CircuitBreaker {
        constructor(name) {
            this.name = name; this.state = CircuitBreakerState.CLOSED;
            this.failureCount = 0; this.successCount = 0;
            this.nextAttemptTime = null; this.lastError = null; this.lastStateChange = Date.now();
        }

        canExecute() {
            if (this.state === CircuitBreakerState.CLOSED) return true;
            if (this.state === CircuitBreakerState.OPEN) {
                if (Date.now() >= this.nextAttemptTime) { this._setState(CircuitBreakerState.HALF_OPEN); this.successCount = 0; return true; }
                throw ErrorHandler.create(ErrorType.CIRCUIT_BREAKER, Messages.ERRORS.CIRCUIT_BREAKER_OPEN, { circuit: this.name, nextAttempt: this.nextAttemptTime });
            }
            return true;
        }

        recordSuccess() {
            this.failureCount = 0; this.lastError = null;
            if (this.state === CircuitBreakerState.HALF_OPEN) {
                this.successCount++;
                if (this.successCount >= CONFIG.CIRCUIT_BREAKER.SUCCESS_THRESHOLD) { this._setState(CircuitBreakerState.CLOSED); Logger.info(`Circuit ${this.name} closed after ${this.successCount} successes`); }
            }
        }

        recordFailure(error = null) {
            this.failureCount++; this.lastError = error;
            if (this.state === CircuitBreakerState.HALF_OPEN) {
                this._setState(CircuitBreakerState.OPEN);
                this.nextAttemptTime = Date.now() + CONFIG.CIRCUIT_BREAKER.RESET_TIMEOUT;
                Logger.warn(`Circuit ${this.name} re-opened from half-open`, { failureCount: this.failureCount });
            } else if (this.failureCount >= CONFIG.CIRCUIT_BREAKER.FAILURE_THRESHOLD) {
                this._setState(CircuitBreakerState.OPEN);
                this.nextAttemptTime = Date.now() + CONFIG.CIRCUIT_BREAKER.RESET_TIMEOUT;
                Logger.warn(`Circuit ${this.name} opened`, { failureCount: this.failureCount });
            }
        }

        reset() { this._setState(CircuitBreakerState.CLOSED); this.failureCount = 0; this.successCount = 0; this.nextAttemptTime = null; this.lastError = null; }

        getState() {
            return { name: this.name, state: this.state, failureCount: this.failureCount, successCount: this.successCount,
                nextAttemptTime: this.nextAttemptTime, lastError: this.lastError?.message || null,
                lastStateChange: this.lastStateChange, timeSinceStateChange: Date.now() - this.lastStateChange };
        }

        _setState(newState) { this.state = newState; this.lastStateChange = Date.now(); }
    }

    const circuitBreakers = { smc: new CircuitBreaker('SMC'), fmc: new CircuitBreaker('FMC'), execution: new CircuitBreaker('Execution') };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 16: CENTRALIZED ERROR HANDLER
     * ═══════════════════════════════════════════════════════════════════════════ */

    const ErrorHandler = {
        create(type, message, context = {}) {
            const error = new Error(message);
            error.type = type; error.context = context; error.timestamp = Date.now(); error.id = Helpers.generateId('err');
            return error;
        },

        handle(error, fallback = null, showToast = true) {
            const errorType = error.type || ErrorType.UNKNOWN;
            const context = error.context || {};
            Logger.error(error.message, { type: errorType, context, stack: error.stack, id: error.id });
            if (showToast && typeof UIController !== 'undefined' && UIController.showToast) {
                const userMessage = this._getUserMessage(errorType, error.message);
                UIController.showToast(userMessage, 'error');
            }
            return fallback;
        },

        async wrap(asyncFn, errorType, fallback = null, showToast = true) {
            try { return await asyncFn(); }
            catch (error) {
                if (error.type) return this.handle(error, fallback, showToast);
                const typedError = this.create(errorType, error.message, { originalError: error.stack });
                return this.handle(typedError, fallback, showToast);
            }
        },

        isRateLimitError(error) {
            const msg = error?.message?.toLowerCase() || '';
            return error?.type === ErrorType.RATE_LIMIT || msg.includes('429') || msg.includes('rate') || msg.includes('throttle') || msg.includes('too many');
        },

        isRetryableError(error) {
            return error?.type === ErrorType.NETWORK || error?.type === ErrorType.TIMEOUT || error?.type === ErrorType.RATE_LIMIT || this.isRateLimitError(error);
        },

        _getUserMessage(errorType, originalMessage) {
            const messages = {
                [ErrorType.NETWORK]: Messages.ERRORS.NETWORK_ERROR, [ErrorType.AUTH]: Messages.ERRORS.AUTH_ERROR,
                [ErrorType.VALIDATION]: originalMessage, [ErrorType.BUSINESS]: originalMessage,
                [ErrorType.TIMEOUT]: Messages.ERRORS.TIMEOUT_ERROR, [ErrorType.PARSE]: Messages.ERRORS.PARSE_ERROR,
                [ErrorType.RATE_LIMIT]: Messages.ERRORS.RATE_LIMITED, [ErrorType.CIRCUIT_BREAKER]: Messages.ERRORS.CIRCUIT_BREAKER_OPEN,
                [ErrorType.SOW]: originalMessage, [ErrorType.UNKNOWN]: Messages.ERRORS.UNKNOWN_ERROR
            };
            return messages[errorType] || messages[ErrorType.UNKNOWN];
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 17: GM REQUEST WRAPPER (Fix #2: added VERSION_CONFLICT handling)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const GMRequest = {
        async fetch(options) {
            const startTime = performance.now();
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: options.method || 'GET', url: options.url, headers: options.headers || {},
                    data: options.body || null, responseType: options.responseType || 'json',
                    withCredentials: true, anonymous: false, timeout: options.timeout || CONFIG.API.REQUEST_TIMEOUT,
                    onload: (response) => {
                        const duration = performance.now() - startTime;
                        if (response.status >= 200 && response.status < 300) {
                            PerformanceMonitor.recordApiCall(duration, false);
                            let data = response.response;
                            if (typeof data === 'string' && options.responseType === 'json') {
                                try { data = JSON.parse(data); }
                                catch (e) { PerformanceMonitor.recordApiCall(duration, true); reject(ErrorHandler.create(ErrorType.PARSE, `JSON parse failed: ${e.message}`, { url: options.url })); return; }
                            }
                            resolve(data);
                        } else if (response.status === 401 || response.status === 403) {
                            PerformanceMonitor.recordApiCall(duration, true);
                            if (response.status === 403) TokenManager.clear();
                            reject(ErrorHandler.create(ErrorType.AUTH, Messages.ERRORS.AUTH_ERROR, { status: response.status }));
                        } else if (response.status === 404) {
                            PerformanceMonitor.recordApiCall(duration, true);
                            reject(ErrorHandler.create(ErrorType.BUSINESS, Messages.ERRORS.RESOURCE_NOT_FOUND, { status: 404 }));
                        } else if (response.status === 409) {
                            // Fix #2: handle version conflict as typed error
                            PerformanceMonitor.recordApiCall(duration, true);
                            reject(ErrorHandler.create(ErrorType.BUSINESS, Messages.ERRORS.VERSION_CONFLICT, { status: 409 }));
                        } else if (response.status === 429) {
                            PerformanceMonitor.recordApiCall(duration, true);
                            reject(ErrorHandler.create(ErrorType.RATE_LIMIT, Messages.ERRORS.RATE_LIMITED, { status: 429 }));
                        } else {
                            PerformanceMonitor.recordApiCall(duration, true);
                            reject(ErrorHandler.create(ErrorType.NETWORK, `HTTP ${response.status}: ${response.statusText}`, { status: response.status }));
                        }
                    },
                    onerror: () => { PerformanceMonitor.recordApiCall(performance.now() - startTime, true); reject(ErrorHandler.create(ErrorType.NETWORK, Messages.ERRORS.NETWORK_ERROR, { url: options.url })); },
                    ontimeout: () => { PerformanceMonitor.recordApiCall(performance.now() - startTime, true); reject(ErrorHandler.create(ErrorType.TIMEOUT, Messages.ERRORS.TIMEOUT_ERROR, { url: options.url })); }
                });
            });
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 18: HTTP CLIENT WITH RETRY
     * ═══════════════════════════════════════════════════════════════════════════ */

    const HttpClient = {
        async request(options, context = 'API call', circuitBreaker = null) {
            if (circuitBreaker) circuitBreaker.canExecute();
            const maxRetries = CONFIG.API.MAX_RETRIES;
            const baseDelay = CONFIG.API.RETRY_BASE_DELAY;

            for (let attempt = 0; attempt <= maxRetries; attempt++) {
                try {
                    const result = await GMRequest.fetch(options);
                    if (circuitBreaker) circuitBreaker.recordSuccess();
                    return result;
                } catch (error) {
                    if (circuitBreaker) circuitBreaker.recordFailure(error);
                    const isLastAttempt = attempt === maxRetries;
                    const isRetryable = ErrorHandler.isRetryableError(error);
                    if (isLastAttempt || !isRetryable) {
                        Logger.error(`${context}: Failed after ${attempt + 1} attempts`, { error: error.message, type: error.type, context: error.context });
                        throw error;
                    }
                    let delay = Math.min(baseDelay * Math.pow(2, attempt) + Math.random() * 200, CONFIG.API.RETRY_MAX_DELAY);
                    if (ErrorHandler.isRateLimitError(error)) { delay *= CONFIG.API.RATE_LIMIT_MULTIPLIER; Logger.warn(`${context}: Rate limited, waiting ${Math.round(delay)}ms...`); }
                    else Logger.warn(`${context}: Attempt ${attempt + 1} failed, retrying in ${Math.round(delay)}ms...`);
                    await sleep(delay);
                }
            }
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 19: TOKEN MANAGER (Fix #1: race condition on fetchPromise)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const TokenManager = (() => {
        let token = null;
        let tokenTime = 0;
        let isFetching = false;
        let fetchPromise = null;
        let autoRefreshInterval = null;

        return {
            init() {
                if (Helpers.isOnSMC()) {
                    const pageToken = this.extractFromPage();
                    if (pageToken) { this.save(pageToken); Logger.info('Token: Extracted from current page'); return true; }
                }
                return this.loadFromStorage();
            },

            extractFromPage() {
                const selectors = ['meta[name="x-csrf-token"]', 'meta[name="csrf-token"]'];
                for (const selector of selectors) { const meta = document.querySelector(selector); if (meta?.content) return meta.content; }
                return null;
            },

            loadFromStorage() {
                try {
                    const saved = GM_getValue(CONFIG.TOKEN.STORAGE_KEY, null);
                    const time = GM_getValue(CONFIG.TOKEN.TIME_KEY, 0);
                    if (saved && time) { token = saved; tokenTime = time; Logger.info('Token: Loaded from storage', { age: this.getAgeString() }); return true; }
                } catch (e) { Logger.error('Token: Load error', e.message); }
                return false;
            },

            save(newToken) {
                token = newToken; tokenTime = Date.now();
                try { GM_setValue(CONFIG.TOKEN.STORAGE_KEY, token); GM_setValue(CONFIG.TOKEN.TIME_KEY, tokenTime); }
                catch (e) { Logger.error('Token: Save error', e.message); }
            },

            isExpired() { if (!token || !tokenTime) return true; return (Date.now() - tokenTime) > CONFIG.TOKEN.MAX_AGE; },

            getRemainingSeconds() {
                if (!token || !tokenTime) return 0;
                const elapsed = Date.now() - tokenTime;
                return Math.ceil(Math.max(0, CONFIG.TOKEN.MAX_AGE - elapsed) / 1000);
            },

            getAgeString() {
                if (!tokenTime) return 'N/A';
                const age = Date.now() - tokenTime;
                if (age < 1000) return 'Fresh';
                if (age < 60000) return `${Math.floor(age / 1000)}s`;
                if (age < 3600000) return `${Math.floor(age / 60000)}m ${Math.floor((age % 60000) / 1000)}s`;
                return `${Math.floor(age / 3600000)}h`;
            },

            getStatus() {
                const remainingSeconds = this.getRemainingSeconds();
                if (isFetching) return { status: 'fetching', text: '⏳', class: 'fetching', remainingSeconds: 0 };
                if (!token) return { status: 'missing', text: '❌', class: 'error', remainingSeconds: 0 };
                if (this.isExpired()) return { status: 'expired', text: '❌', class: 'error', remainingSeconds: 0 };
                if (remainingSeconds <= CONFIG.TOKEN.CRITICAL_THRESHOLD) return { status: 'critical', text: `⚠️ ${remainingSeconds}s`, class: 'critical', remainingSeconds };
                if (remainingSeconds <= CONFIG.TOKEN.WARNING_THRESHOLD) return { status: 'warning', text: `✓ ${remainingSeconds}s`, class: 'warning', remainingSeconds };
                return { status: 'ready', text: `✓ ${remainingSeconds}s`, class: 'ready', remainingSeconds };
            },

            getToken() { return token; },

            startAutoRefresh() {
                this.stopAutoRefresh();
                autoRefreshInterval = setInterval(async () => {
                    const remaining = this.getRemainingSeconds();
                    if (remaining < CONFIG.TOKEN.WARNING_THRESHOLD) { Logger.info('Token: Auto-refreshing...'); await this.ensure(); }
                }, CONFIG.TOKEN.REFRESH_INTERVAL);
                Logger.debug('Token: Auto-refresh started');
            },

            stopAutoRefresh() { if (autoRefreshInterval) { clearInterval(autoRefreshInterval); autoRefreshInterval = null; Logger.debug('Token: Auto-refresh stopped'); } },

            // Fix #1: Removed fetchLock. Gate solely on fetchPromise.
            async ensure() {
                if (Helpers.isOnSMC()) {
                    const pageToken = this.extractFromPage();
                    if (pageToken) { this.save(pageToken); Logger.debug('Token: Refreshed from page'); return true; }
                }
                if (token && !this.isExpired()) { Logger.debug('Token: Valid, no refresh needed'); return true; }
                // If a fetch is already in flight, wait for it
                if (fetchPromise) {
                    Logger.debug('Token: Waiting for existing fetch...');
                    return await fetchPromise;
                }
                isFetching = true;
                Logger.info('Token: Fetching from SMC...');
                fetchPromise = this._doFetch();
                try {
                    return await fetchPromise;
                } finally {
                    isFetching = false;
                    fetchPromise = null;
                }
            },

            _doFetch() {
                return new Promise((resolve) => {
                    GM_xmlhttpRequest({
                        method: 'GET', url: CONFIG.URLS.SMC_BASE,
                        headers: { 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8' },
                        withCredentials: true, timeout: CONFIG.TOKEN.FETCH_TIMEOUT,
                        onload: (response) => {
                            Logger.debug('Token fetch response', { status: response.status });
                            if (response.status === 200) {
                                const patterns = [
                                    /<meta[^>]+name=["']x-csrf-token["'][^>]+content=["']([^"']+)["']/i,
                                    /<meta[^>]+content=["']([^"']+)["'][^>]+name=["']x-csrf-token["']/i
                                ];
                                let foundToken = null;
                                for (const pattern of patterns) { const match = response.responseText.match(pattern); if (match?.[1]) { foundToken = match[1]; break; } }
                                if (foundToken) {
                                    this.save(foundToken);
                                    Logger.info('Token: Fetched successfully');
                                    Telemetry.track(TelemetryEventType.TOKEN_REFRESH, { success: true });
                                    resolve(true);
                                } else {
                                    Logger.warn('Token: Not found in response');
                                    Telemetry.track(TelemetryEventType.TOKEN_REFRESH, { success: false, reason: 'not_found' });
                                    resolve(false);
                                }
                            } else {
                                Logger.warn('Token fetch failed', { status: response.status });
                                Telemetry.track(TelemetryEventType.TOKEN_REFRESH, { success: false, status: response.status });
                                resolve(false);
                            }
                        },
                        onerror: (error) => { Logger.error('Token: Network error', error); Telemetry.track(TelemetryEventType.TOKEN_REFRESH, { success: false, reason: 'network_error' }); resolve(false); },
                        ontimeout: () => { Logger.error('Token: Timeout'); Telemetry.track(TelemetryEventType.TOKEN_REFRESH, { success: false, reason: 'timeout' }); resolve(false); }
                    });
                });
            },

            clear() {
                token = null; tokenTime = 0; this.stopAutoRefresh();
                try { GM_setValue(CONFIG.TOKEN.STORAGE_KEY, null); GM_setValue(CONFIG.TOKEN.TIME_KEY, 0); }
                catch (e) { Logger.error('Token: Clear error', e.message); }
            },

            cleanup() { this.stopAutoRefresh(); }
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 20: SOW CONFIG MANAGER
     * ═══════════════════════════════════════════════════════════════════════════ */

    const SOWConfigManager = (() => {
        let configMap = new Map();
        let allShippersData = [];
        let cachedShipperNames = null;
        let lastFetchTime = 0;
        let isLoading = false;
        let loadingPromise = null;

        const REQUIRED_FIELDS = [
            'Title', 'Rate', 'RateUnit', 'MaxCharge',
            'PU_Live_Eligible', 'PU_Live_FreeTime', 'PU_DropHook_Eligible', 'PU_DropHook_FreeTime',
            'DO_Live_Eligible', 'DO_Live_FreeTime', 'DO_DropHook_Eligible', 'DO_DropHook_FreeTime'
        ];

        const validateRequiredFields = (item) => {
            const errors = []; const missingFields = [];
            for (const field of REQUIRED_FIELDS) {
                const value = item[field];
                if (value === null || value === undefined || value === '') { errors.push(`Missing: ${field}`); missingFields.push(field); continue; }
                if (field === 'Title' && (typeof value !== 'string' || value.trim() === '')) errors.push(`${field}: Must be non-empty text`);
                if (field === 'Rate') { const numVal = parseFloat(value); if (isNaN(numVal) || numVal < 0) errors.push(`${field}: Must be a positive number`); }
                if (field === 'RateUnit') { const normalized = String(value).toUpperCase(); if (!['HOUR', 'MINUTE'].includes(normalized)) errors.push(`${field}: Must be "Hour" or "Minute", got "${value}"`); }
                if (field === 'MaxCharge') { const numVal = parseFloat(value); if (isNaN(numVal) || numVal < 0) errors.push(`${field}: Must be a positive number`); }
                if (field.includes('FreeTime')) { const numVal = parseFloat(value); if (isNaN(numVal) || numVal < 0) errors.push(`${field}: Must be a non-negative number`); }
                if (field.includes('Eligible')) { const boolVal = parseBoolean(value, null); if (boolVal === null) errors.push(`${field}: Must be Yes or No`); }
            }
            return { isValid: errors.length === 0, errors, missingFields };
        };

        const parseResponse = (response) => {
            const validMap = new Map(); const allShippers = [];
            if (!response?.d?.results || !Array.isArray(response.d.results)) throw ErrorHandler.create(ErrorType.PARSE, 'Invalid SharePoint response structure');
            const items = response.d.results;
            Logger.info(`SOW: Parsing ${items.length} shipper configurations`);

            for (const item of items) {
                const shipperName = item.Title || 'Unknown';
                const validation = validateRequiredFields(item);
                const isActive = parseBoolean(item.IsActive, true);
                let status = ShipperStatus.ACTIVE;
                if (!validation.isValid) status = ShipperStatus.VALIDATION_ERROR;
                else if (!isActive) status = ShipperStatus.INACTIVE;

                const shipperData = {
                    shipperName, status, isActive,
                    validationErrors: validation.errors, missingFields: validation.missingFields,
                    rawData: {
                        Title: item.Title, Rate: item.Rate, RateUnit: item.RateUnit, MaxCharge: item.MaxCharge,
                        BillingIncrement: item.BillingIncrement, RoundingRule: item.RoundingRule, RoundDownMaxMinutes: item.RoundDownMaxMinutes,
                        RequiresApproval: item.RequiresApproval, AutoChargeAllowed: item.AutoChargeAllowed, AuthNumberRequired: item.AuthNumberRequired,
                        IsActive: item.IsActive, Notes: item.Notes,
                        PU_Live_Eligible: item.PU_Live_Eligible, PU_Live_FreeTime: item.PU_Live_FreeTime,
                        PU_DropHook_Eligible: item.PU_DropHook_Eligible, PU_DropHook_FreeTime: item.PU_DropHook_FreeTime,
                        DO_Live_Eligible: item.DO_Live_Eligible, DO_Live_FreeTime: item.DO_Live_FreeTime,
                        DO_DropHook_Eligible: item.DO_DropHook_Eligible, DO_DropHook_FreeTime: item.DO_DropHook_FreeTime
                    }
                };
                allShippers.push(shipperData);

                if (validation.isValid) {
                    const rateUnit = String(item.RateUnit || 'Hour').toUpperCase();
                    const config = {
                        shipperName, rate: parseFloat(item.Rate), rateUnit: rateUnit === 'MINUTE' ? 'MINUTE' : 'HOUR',
                        maxCharge: parseFloat(item.MaxCharge),
                        billingIncrement: item.BillingIncrement ? parseFloat(item.BillingIncrement) : null,
                        roundingRule: item.RoundingRule ? String(item.RoundingRule).toUpperCase() : null,
                        roundDownMaxMinutes: item.RoundDownMaxMinutes ? parseFloat(item.RoundDownMaxMinutes) : null,
                        requiresApproval: parseBoolean(item.RequiresApproval, false),
                        autoChargeAllowed: parseBoolean(item.AutoChargeAllowed, false),
                        authNumberRequired: parseBoolean(item.AuthNumberRequired, false),
                        isActive, notes: item.Notes || '', isComplete: true,
                        rules: {
                            PICKUP: {
                                LIVE: { eligible: parseBoolean(item.PU_Live_Eligible, false), freeTime: parseFloat(item.PU_Live_FreeTime) || 0 },
                                DROP_HOOK: { eligible: parseBoolean(item.PU_DropHook_Eligible, false), freeTime: parseFloat(item.PU_DropHook_FreeTime) || 0 }
                            },
                            DROP_OFF: {
                                LIVE: { eligible: parseBoolean(item.DO_Live_Eligible, false), freeTime: parseFloat(item.DO_Live_FreeTime) || 0 },
                                DROP_HOOK: { eligible: parseBoolean(item.DO_DropHook_Eligible, false), freeTime: parseFloat(item.DO_DropHook_FreeTime) || 0 }
                            }
                        },
                        displayInfo: {
                            rate: rateUnit === 'MINUTE' ? `$${item.Rate}/min` : `$${item.Rate}/hr`,
                            max: `$${item.MaxCharge}`,
                            billingIncrement: item.BillingIncrement ? `${item.BillingIncrement} min` : '-',
                            puLiveFreeTime: `${item.PU_Live_FreeTime} min`, puDropHookFreeTime: `${item.PU_DropHook_FreeTime} min`,
                            doLiveFreeTime: `${item.DO_Live_FreeTime} min`, doDropHookFreeTime: `${item.DO_DropHook_FreeTime} min`
                        }
                    };
                    validMap.set(shipperName, config);
                    Logger.debug(`SOW: ✓ Loaded "${shipperName}"`, { status });
                } else {
                    Logger.warn(`SOW: ✗ Validation failed for "${shipperName}"`, { errors: validation.errors });
                }
            }
            return { validMap, allShippers };
        };

        return {
            async init() { return this.fetch(); },

            async fetch() {
                if (isLoading && loadingPromise) { Logger.debug('SOW: Already loading, waiting for existing request'); return loadingPromise; }
                isLoading = true;
                AppState.set('sowStatus', SOWStatus.LOADING);
                if (typeof UIController !== 'undefined' && UIController.updateSOWIndicator) UIController.updateSOWIndicator();
                Logger.info('SOW: Fetching configuration from SharePoint...');
                loadingPromise = this._doFetch();
                try { return await loadingPromise; } finally { isLoading = false; loadingPromise = null; }
            },

            async _doFetch() {
                try {
                    const response = await GMRequest.fetch({
                        method: 'GET', url: CONFIG.SHAREPOINT.API_URL,
                        headers: { 'Accept': 'application/json;odata=verbose', 'Content-Type': 'application/json' },
                        timeout: CONFIG.SHAREPOINT.REQUEST_TIMEOUT
                    });
                    const { validMap, allShippers } = parseResponse(response);
                    configMap = validMap; allShippersData = allShippers;
                    cachedShipperNames = null;
                    lastFetchTime = Date.now();
                    AppState.update({ sowStatus: SOWStatus.LOADED, sowShipperCount: configMap.size, sowLastError: null, sowLastRefresh: lastFetchTime });
                    Logger.info(`SOW: ✓ Successfully loaded ${configMap.size} valid shipper configurations (${allShippers.length} total)`);
                    Telemetry.track(TelemetryEventType.SOW_LOAD, { success: true, validCount: configMap.size, totalCount: allShippers.length });
                    if (typeof UIController !== 'undefined') { UIController.updateSOWIndicator(); UIController.showToast(Messages.SUCCESS.SOW_LOADED(configMap.size), 'success'); }
                    return true;
                } catch (error) {
                    Logger.error('SOW: Failed to fetch configuration', error.message);
                    let status = SOWStatus.ERROR; let errorMessage = Messages.ERRORS.SOW_SERVER_UNREACHABLE;
                    if (error.type === ErrorType.AUTH || error.message?.includes('401') || error.message?.includes('403')) { status = SOWStatus.AUTH_REQUIRED; errorMessage = Messages.ERRORS.SOW_AUTH_REQUIRED; }
                    AppState.update({ sowStatus: status, sowShipperCount: 0, sowLastError: errorMessage });
                    Telemetry.track(TelemetryEventType.SOW_LOAD, { success: false, error: error.message });
                    if (typeof UIController !== 'undefined') { UIController.updateSOWIndicator(); UIController.showToast(errorMessage, 'error'); }
                    return false;
                }
            },

            getConfig(shipperName) {
                if (!shipperName) { Logger.warn('SOW: getConfig called with empty shipper name'); return null; }
                const config = configMap.get(shipperName);
                if (!config) Logger.warn(`SOW: No configuration found for "${shipperName}"`);
                return config;
            },

            validateShipper(shipperName) {
                const result = { valid: false, config: null, error: null, errorType: null };
                if (AppState.get('sowStatus') !== SOWStatus.LOADED) { result.error = Messages.ERRORS.SOW_SERVER_UNREACHABLE; result.errorType = ResultType.SOW_NOT_CONFIGURED; return result; }
                const config = this.getConfig(shipperName);
                if (!config) { result.error = Messages.ERRORS.SOW_NOT_CONFIGURED(shipperName); result.errorType = ResultType.SOW_NOT_CONFIGURED; return result; }
                if (!config.isActive) { result.error = Messages.ERRORS.SOW_DISABLED(shipperName); result.errorType = ResultType.SOW_DISABLED; return result; }
                if (!config.isComplete) { result.error = Messages.ERRORS.SOW_INCOMPLETE(shipperName); result.errorType = ResultType.SOW_INCOMPLETE; return result; }
                result.valid = true; result.config = config;
                return result;
            },

            getShipperNames() {
                if (!cachedShipperNames) cachedShipperNames = Array.from(configMap.keys());
                return cachedShipperNames;
            },

            getAllShippersData() { return [...allShippersData]; },

            getStatistics() {
                const stats = { total: allShippersData.length, active: 0, inactive: 0, validationErrors: 0, hourlyRate: 0, minuteRate: 0 };
                for (const shipper of allShippersData) {
                    if (shipper.status === ShipperStatus.ACTIVE) {
                        stats.active++; const rateUnit = String(shipper.rawData.RateUnit || '').toUpperCase();
                        if (rateUnit === 'MINUTE') stats.minuteRate++; else stats.hourlyRate++;
                    } else if (shipper.status === ShipperStatus.INACTIVE) stats.inactive++;
                    else if (shipper.status === ShipperStatus.VALIDATION_ERROR) stats.validationErrors++;
                }
                return stats;
            },

            isLoaded() { return AppState.get('sowStatus') === SOWStatus.LOADED && configMap.size > 0; },
            getShipperCount() { return configMap.size; },
            getLastRefreshTime() { return lastFetchTime; },

            clear() {
                configMap.clear(); allShippersData = []; cachedShipperNames = null; lastFetchTime = 0;
                AppState.update({ sowStatus: SOWStatus.NOT_LOADED, sowShipperCount: 0 });
            },

            getStatus() {
                return { status: AppState.get('sowStatus'), shipperCount: configMap.size, totalShippers: allShippersData.length, lastFetchTime, lastError: AppState.get('sowLastError'), isLoading };
            }
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 21: DATA HELPERS
     * ═══════════════════════════════════════════════════════════════════════════ */

    const DataHelpers = {
        formatStatusDisplay(statusCode) {
            if (!statusCode) return { display: 'Unknown', color: '#888888', group: 'unknown' };
            if (OrderStatusMap[statusCode]) return OrderStatusMap[statusCode];
            const upperStatus = statusCode.toUpperCase();
            let group = 'active'; let color = '#f0ad4e';
            if (upperStatus.includes('CANCEL') || upperStatus.includes('REJECT')) { group = 'cancelled'; color = '#d9534f'; }
            else if (upperStatus.includes('PAID')) { group = 'paid'; color = '#5cb85c'; }
            else if (upperStatus.includes('INVOICE') || upperStatus === 'PENDING_PAYMENT') { group = 'invoiced'; color = '#5cb85c'; }
            const display = statusCode.split('_').map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join(' ');
            return { display, color, group };
        },

        formatLoadType(loadingType) {
            if (!loadingType) return { display: 'Unknown', icon: '📦', isDropHook: false, key: 'LIVE' };
            const upper = loadingType.toUpperCase();
            const isDropHook = upper.includes('DROP') || upper.includes('HOOK');
            return { display: isDropHook ? 'Drop Hook' : 'Live', icon: isDropHook ? '🪝' : '🔄', isDropHook, key: isDropHook ? 'DROP_HOOK' : 'LIVE' };
        },

        formatStopType(stopActionType) {
            if (!stopActionType) return { display: 'Unknown', label: 'STOP', class: 'pickup', isPickup: false, key: 'DROP_OFF' };
            const upper = stopActionType.toUpperCase();
            const isPickup = upper.includes('PICKUP') || upper === 'PICK_UP';
            return { display: isPickup ? 'Pickup' : 'Drop Off', label: isPickup ? 'SHIPPER' : 'RECEIVER', class: isPickup ? 'pickup' : 'dropoff', isPickup, key: isPickup ? 'PICKUP' : 'DROP_OFF' };
        },

        findDetentionHolds(pricing) {
            const result = { shipper: false, receiver: false, shipperCode: null, receiverCode: null, shipperItem: null, receiverItem: null, shipperAmount: 0, receiverAmount: 0, shipperExists: false, receiverExists: false };
            if (!Array.isArray(pricing)) return result;
            const DRIVER_DETENTION_SHIPPER = 'DETENTION_DRIVER_AT_SHIPPER';
            const DRIVER_DETENTION_RECEIVER = 'DETENTION_DRIVER_AT_RECEIVER';
            for (let i = 0; i < pricing.length; i++) {
                const item = pricing[i]; const code = String(item.pricingCode || '').toUpperCase(); const value = item.price?.value || 0;
                if (code === DRIVER_DETENTION_SHIPPER) { result.shipperCode = item.pricingCode; result.shipperItem = item; result.shipperAmount = value; result.shipperExists = true; if (value === 0) result.shipper = true; }
                if (code === DRIVER_DETENTION_RECEIVER) { result.receiverCode = item.pricingCode; result.receiverItem = item; result.receiverAmount = value; result.receiverExists = true; if (value === 0) result.receiver = true; }
            }
            return result;
        },

        calculateTimeDiff(planned, actual) {
            if (!planned || !actual) return { minutes: null, text: 'Pending', class: 'pending', label: '', status: 'UNKNOWN' };
            try {
                const plannedDate = typeof planned === 'number' ? planned : new Date(planned).getTime();
                const actualDate = typeof actual === 'number' ? actual : new Date(actual).getTime();
                if (isNaN(plannedDate) || isNaN(actualDate)) return { minutes: null, text: 'Invalid', class: 'pending', label: '', status: 'UNKNOWN' };
                const plannedMinutes = Math.floor(plannedDate / 60000) * 60000;
                const actualMinutes = Math.floor(actualDate / 60000) * 60000;
                const diffMinutes = Math.round((actualMinutes - plannedMinutes) / 60000);
                let status;
                if (diffMinutes <= CONFIG.TIMING.EARLY_MINUTES) status = 'EARLY';
                else if (diffMinutes <= CONFIG.TIMING.ON_TIME_MINUTES) status = 'ON_TIME';
                else status = 'LATE';
                if (diffMinutes > 0) return { minutes: diffMinutes, text: Helpers.formatDuration(diffMinutes), class: 'late', label: 'LATE', status };
                if (diffMinutes < 0) return { minutes: diffMinutes, text: Helpers.formatDuration(Math.abs(diffMinutes)), class: 'early', label: 'EARLY', status };
                return { minutes: 0, text: 'ON TIME', class: 'on-time', label: '', status: 'ON_TIME' };
            } catch (e) { Logger.warn('calculateTimeDiff error', e.message); return { minutes: null, text: 'Error', class: 'pending', label: '', status: 'UNKNOWN' }; }
        },

        getDetentionPricingConfig(isPickup) { return isPickup ? DetentionPricing.SHIPPER : DetentionPricing.RECEIVER; },

        getActionDisplayText(analysis) {
            if (!analysis) return Messages.INFO.NO_ACTION_NEEDED;

            // Lane mismatch
            if (analysis.laneMismatch) {
                if (analysis.hasHold) return '🚫 Lane Mismatch — Hold Released';
                return '🚫 Lane Mismatch — No charge';
            }

            if (analysis.mismatchType) {
                const statusTag = analysis.orderStatusGroup === 'invoiced' ? ' [Invoiced]' : analysis.orderStatusGroup === 'paid' ? ' [Paid]' : '';
                switch (analysis.mismatchType) {
                    case 'UNDER':
                        return `⚠️ Undercharged: SMC $${analysis.existingAmount.toFixed(2)} / Correct $${analysis.calculatedAmount.toFixed(2)}${statusTag}`;
                    case 'OVER':
                        return `⚠️ Overcharged: SMC $${analysis.existingAmount.toFixed(2)} / Correct $${analysis.calculatedAmount.toFixed(2)}${statusTag}`;
                    case 'SHOULD_NOT_EXIST':
                        return `🚨 Should Not Be Charged: SMC $${analysis.existingAmount.toFixed(2)} / Correct $0.00 (${analysis.noChargeReason || 'N/A'})${statusTag}`;
                    case 'CORRECT':
                        return `✅ Charge Correct ($${analysis.existingAmount.toFixed(2)})${statusTag}`;
                }
            }

            if (analysis.processed) {
                switch (analysis.processedAction) {
                    case 'updated': return `$${analysis.processedAmount.toFixed(2)} (Charge Added)`;
                    case 'created': return `$${analysis.processedAmount.toFixed(2)} (Recovered)`;
                    case 'released': return '$0 (Hold Released)';
                    case 'released_lane_mismatch': return '$0 (Lane Mismatch — Released)';
                    case 'hold_created': return `$0 Hold Created ($${analysis.charge.toFixed(2)} chargeable)`;
                    case 'skipped': return 'Skipped';
                    case 'timeout': return 'Timeout';
                    case 'analysis_only': return `$${analysis.charge.toFixed(2)} (Analysis Only)`;
                    default: return Messages.INFO.NO_ACTION_NEEDED;
                }
            }

            let text = '';
            switch (analysis.type) {
                case ResultType.LANE_MISMATCH:
                    if (analysis.hasHold) text = '$0 (Lane Mismatch — Release Hold)';
                    else text = '$0 (Lane Mismatch — No charge)';
                    break;
                case ResultType.CHARGEABLE: case ResultType.CHARGEABLE_MAX:
                    if (analysis.action === ActionType.ANALYSIS_ONLY) text = `$${analysis.charge.toFixed(2)} (Analysis Only - Hold Exists)`;
                    else if (analysis.action === ActionType.CREATE_HOLD_ONLY) text = `$${analysis.charge.toFixed(2)} (Will Create $0 Hold)`;
                    else if (analysis.action === ActionType.PENDING_APPROVAL) text = `$${analysis.charge.toFixed(2)} (Approval Required)`;
                    else text = `$${analysis.charge.toFixed(2)} (Chargeable)`;
                    break;
                case ResultType.CHARGE_EXISTS: text = `$${analysis.existingAmount.toFixed(2)} (Already Exists)`; break;
                case ResultType.WITHIN_FREE_TIME: case ResultType.NO_HOLD_NO_CHARGE: text = '$0 (Within Free Time)'; break;
                case ResultType.BELOW_MINIMUM_THRESHOLD: text = '$0 (Below Minimum)'; break;
                case ResultType.DRIVER_LATE: text = '$0 (Driver Late)'; break;
                case ResultType.NO_DETENTION_DROP_HOOK: text = '$0 (Drop & Hook)'; break;
                case ResultType.MISSING_ARRIVAL: text = 'Awaiting Arrival'; break;
                case ResultType.MISSING_DEPARTURE: text = 'Awaiting Departure'; break;
                case ResultType.ORDER_CANCELLED: text = 'Order Cancelled'; break;
                case ResultType.ORDER_INVOICED: text = 'Already Invoiced'; break;
                case ResultType.FMC_DATA_UNAVAILABLE: text = 'FMC Unavailable'; break;
                case ResultType.SOW_NOT_CONFIGURED: text = 'SOW Not Configured'; break;
                case ResultType.SOW_DISABLED: text = 'SOW Disabled'; break;
                case ResultType.SOW_INCOMPLETE: text = 'SOW Incomplete'; break;
                default: text = Messages.INFO.NO_ACTION_NEEDED;
            }
            if (analysis.approvalNote) text += ` (${analysis.approvalNote})`;
            return text;
        },

        getBreakdownDetails(analysis, fmcStopData, timezone) {
            const details = [];
            if (!analysis || !fmcStopData?.timestamps) { details.push('• FMC data unavailable'); return details; }

            // Lane mismatch breakdown
            if (analysis.laneMismatch && analysis.laneValidation) {
                const lv = analysis.laneValidation;
                details.push('• 🚫 LANE MISMATCH DETECTED');
                details.push(`• SMC Origin: ${lv.smcOrigin || 'N/A'}`);
                details.push(`• FMC Origin: ${lv.fmcOrigin || 'N/A'} ${lv.originMatch ? '✅' : '❌ MISMATCH'}`);
                details.push(`• SMC Dest: ${lv.smcDest || 'N/A'}`);
                details.push(`• FMC Dest: ${lv.fmcDest || 'N/A'} ${lv.destMatch ? '✅' : '❌ MISMATCH'}`);
                details.push('• Action: Detention not applicable for this VR');
                if (analysis.hasHold) details.push('• Hold released with timestamp comment');
                return details;
            }

            const ts = fmcStopData.timestamps;

            if (ts.actualYardArrival && ts.plannedYardArrival) {
                const arrivalDiff = this.calculateTimeDiff(ts.plannedYardArrival, ts.actualYardArrival);
                const arrivalStatus = arrivalDiff.minutes <= 0 ? 'On time' : 'Late';
                details.push(`• Arrival: ${arrivalStatus} (${Helpers.formatDelta(arrivalDiff.minutes)})`);
            } else if (!ts.actualYardArrival) { details.push('• Arrival: Pending'); }

            if (ts.actualYardDeparture && ts.plannedYardDeparture) {
                const depDiff = this.calculateTimeDiff(ts.plannedYardDeparture, ts.actualYardDeparture);
                const depStatus = depDiff.minutes <= 0 ? 'On time' : `${Helpers.formatDuration(depDiff.minutes)} late`;
                details.push(`• Departure: ${depStatus}`);
            } else if (!ts.actualYardDeparture) { details.push('• Departure: Pending'); }

            if (analysis.hasBorrowedTimestamps) details.push('• ⚠️ Timestamp Issue: Data borrowed from another VR in tour');

            if (analysis.breakdown) {
                const lines = analysis.breakdown.split('\n');
                lines.forEach(line => { if (line.trim()) details.push(`• ${line.trim()}`); });
            }

            if (analysis.mismatchType) {
                const statusTag = analysis.orderStatusGroup === 'invoiced' ? ' [Invoiced]' : analysis.orderStatusGroup === 'paid' ? ' [Paid]' : '';
                switch (analysis.mismatchType) {
                    case 'UNDER':
                        details.push(`• ⚠️ MISMATCH: Current $${analysis.existingAmount.toFixed(2)} → Correct $${analysis.calculatedAmount.toFixed(2)} (Undercharged by $${analysis.mismatchAmount.toFixed(2)})${statusTag}`);
                        break;
                    case 'OVER':
                        details.push(`• ⚠️ MISMATCH: Current $${analysis.existingAmount.toFixed(2)} → Correct $${analysis.calculatedAmount.toFixed(2)} (Overcharged by $${analysis.mismatchAmount.toFixed(2)})${statusTag}`);
                        break;
                    case 'SHOULD_NOT_EXIST':
                        details.push(`• 🚨 SHOULD NOT EXIST: Current $${analysis.existingAmount.toFixed(2)} → Correct $0.00 (${analysis.noChargeReason || 'N/A'})${statusTag}`);
                        break;
                    case 'CORRECT':
                        details.push(`• ✅ Charge is correct ($${analysis.existingAmount.toFixed(2)})${statusTag}`);
                        break;
                }
            }

            if (analysis.processed) {
                switch (analysis.processedAction) {
                    case 'updated': details.push(`• Action: Hold updated to $${analysis.processedAmount.toFixed(2)}`); break;
                    case 'created': details.push(`• Action: New charge created $${analysis.processedAmount.toFixed(2)}`); break;
                    case 'released': details.push('• Action: Hold released ($0)'); break;
                    case 'released_lane_mismatch': details.push('• Action: Hold released (Lane Mismatch)'); break;
                    case 'hold_created': details.push('• Action: $0 hold created for manual review'); break;
                    case 'analysis_only': details.push('• Action: Analysis only (no auto-charge)'); break;
                }
            } else if (analysis.action === ActionType.RELEASE_HOLD) { details.push('• Action: Will release hold'); }
            else if (analysis.action === ActionType.ADD_CHARGE_UPDATE || analysis.action === ActionType.ADD_CHARGE_CREATE) { details.push(`• Action: Will add charge $${analysis.charge.toFixed(2)}`); }
            else if (analysis.action === ActionType.CREATE_HOLD_ONLY) { details.push('• Action: Will create $0 hold for manual review'); }
            else if (analysis.action === ActionType.ANALYSIS_ONLY) { details.push('• Action: Analysis only (auto-charge disabled)'); }
            else if (analysis.action === ActionType.PENDING_APPROVAL) { details.push('• Action: Requires approval'); }
            else if (analysis.type === ResultType.NO_HOLD_NO_CHARGE || analysis.type === ResultType.WITHIN_FREE_TIME) { details.push('• No charge applicable'); }

            if (analysis.approvalNote) details.push(`• ⚠️ ${analysis.approvalNote}`);
            return details;
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 21.5: LANE VALIDATOR (NEW — contracted lane validation)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const LaneValidator = {
        /**
         * Validates that the FMC VR lane matches the SMC contracted lane.
         * @param {Object} viewData - SMC order view data (contains commodities)
         * @param {Object} matchingVR - FMC VR record (contains aggregatedStops)
         * @returns {Object} { isMatch, smcOrigin, smcDest, fmcOrigin, fmcDest, originMatch, destMatch }
         */
        validate(viewData, matchingVR) {
            const result = {
                isMatch: true,
                smcOrigin: null, smcDest: null,
                fmcOrigin: null, fmcDest: null,
                originMatch: true, destMatch: true,
                validated: false
            };

            // Extract SMC contracted lane from commodities
            const commodities = viewData?.commodities;
            if (!commodities || commodities.length === 0) {
                Logger.warn('LaneValidator: No commodities found in viewData');
                return result; // Can't validate — assume match
            }

            const commodity = commodities[0];
            result.smcOrigin = commodity.originStopLocationCode || null;
            result.smcDest = commodity.destinationStopLocationCode || null;

            if (!result.smcOrigin || !result.smcDest) {
                Logger.warn('LaneValidator: Missing stop location codes in commodity', {
                    origin: result.smcOrigin, dest: result.smcDest
                });
                return result; // Can't validate — assume match
            }

            // Extract FMC VR lane from aggregatedStops
            const stops = matchingVR?.aggregatedStops;
            if (!stops || stops.length < 2) {
                Logger.warn('LaneValidator: Insufficient FMC stops for lane validation');
                return result; // Can't validate — assume match
            }

            result.fmcOrigin = stops[0]?.stopCode || null;
            result.fmcDest = stops[stops.length - 1]?.stopCode || null;

            if (!result.fmcOrigin || !result.fmcDest) {
                Logger.warn('LaneValidator: Missing FMC stop codes');
                return result; // Can't validate — assume match
            }

            // Compare — exact string match
            result.originMatch = result.smcOrigin === result.fmcOrigin;
            result.destMatch = result.smcDest === result.fmcDest;
            result.isMatch = result.originMatch && result.destMatch;
            result.validated = true;

            if (!result.isMatch) {
                Logger.info(`LaneValidator: ❌ MISMATCH detected`, {
                    smcLane: `${result.smcOrigin} → ${result.smcDest}`,
                    fmcLane: `${result.fmcOrigin} → ${result.fmcDest}`,
                    originMatch: result.originMatch,
                    destMatch: result.destMatch
                });
            } else {
                Logger.debug(`LaneValidator: ✓ Lane match confirmed`, {
                    lane: `${result.smcOrigin} → ${result.smcDest}`
                });
            }

            return result;
        },

        /**
         * Generates the lane mismatch comment.
         */
        generateComment(laneValidation) {
            return Messages.COMMENTS.LANE_MISMATCH(
                laneValidation.smcOrigin || 'Unknown',
                laneValidation.smcDest || 'Unknown',
                laneValidation.fmcOrigin || 'Unknown',
                laneValidation.fmcDest || 'Unknown'
            );
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 22: FMC API SERVICE
     * ═══════════════════════════════════════════════════════════════════════════ */

    const FMCApiService = {
        async fetchSMCExecution(orderId) {
            const url = `${CONFIG.URLS.SMC_EXECUTION_API}/operator-user-shipment?orderId=${encodeURIComponent(orderId)}`;
            Logger.debug(`FMC API: Fetching SMC execution for ${orderId}`);
            const data = await HttpClient.request({ method: 'GET', url, headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' } }, 'SMC Execution', circuitBreakers.execution);
            return this._parseSMCExecutionResponse(data);
        },

        _parseSMCExecutionResponse(data) {
            if (!data) throw ErrorHandler.create(ErrorType.PARSE, Messages.ERRORS.EMPTY_RESPONSE, { source: 'SMC' });
            if (!data.executionLegs || data.executionLegs.length === 0) throw ErrorHandler.create(ErrorType.BUSINESS, Messages.ERRORS.NO_EXECUTION_LEGS, { data });
            const leg = data.executionLegs[0];
            if (!leg?.tourId) throw ErrorHandler.create(ErrorType.BUSINESS, Messages.ERRORS.NO_TOUR_ID, { leg });
            return {
                orderId: data.orderId, shipmentId: data.shipmentId, shipmentStatus: data.shipmentStatus, shipperId: data.shipperId, freightType: data.freightType,
                tourId: leg.tourId, vehicleRunId: leg.vehicleRunId, carrierId: leg.carrierId, executionStatus: leg.status, transportMode: leg.transportMode,
                origin: { nodeCode: leg.from?.facility?.ncsNodeCode || 'Unknown', facilityName: leg.from?.facility?.facilityName || 'Unknown', timezone: leg.from?.facility?.timezone || 'America/Los_Angeles', loadingType: leg.from?.loadingType || 'Unknown' },
                destination: { nodeCode: leg.to?.facility?.ncsNodeCode || 'Unknown', facilityName: leg.to?.facility?.facilityName || 'Unknown', timezone: leg.to?.facility?.timezone || 'America/Los_Angeles', loadingType: leg.to?.loadingType || 'Unknown' },
                contractedLane: `${leg.from?.facility?.ncsNodeCode || ''}->${leg.to?.facility?.ncsNodeCode || ''}`,
                alerts: data.alerts || []
            };
        },

        async fetchFMCByTourId(tourId) {
            const url = `${CONFIG.URLS.FMC_BASE}/fmc/search/execution/by-id`;
            const requestBody = { searchIds: [tourId], searchByIds: true, page: 0, pageSize: 50, sortOrder: [{ field: "first_dock_arrival_time", dir: "asc" }], bookmarkedSavedSearch: false, executionViewModePreference: "vrs" };
            Logger.debug(`FMC API: Fetching FMC data for Tour ${tourId}`);
            const data = await HttpClient.request({ method: 'POST', url, headers: { 'Accept': 'application/json', 'Content-Type': 'application/json' }, body: JSON.stringify(requestBody) }, 'FMC Data', circuitBreakers.fmc);
            if (!data?.success) throw ErrorHandler.create(ErrorType.BUSINESS, data?.errorMessage || Messages.ERRORS.FMC_API_FAILURE, { tourId });
            Logger.debug('FMC API: FMC data received', { recordCount: data.returnedObject?.records?.length || 0 });
            return data.returnedObject;
        },

        findMatchingVR(fmcRecords, contractedLane) {
            if (!fmcRecords || fmcRecords.length === 0 || !contractedLane) return null;
            const [originCode, destCode] = contractedLane.split('->');
            Logger.debug(`FMC API: Finding VR for lane ${contractedLane}`);
            const strategies = [
                (vr) => vr?.simpleFacilityLane === contractedLane,
                (vr) => vr?.facilityLaneString === contractedLane,
                (vr) => { if (!vr?.aggregatedStops || vr.aggregatedStops.length < 2) return false; const firstStop = vr.aggregatedStops[0]?.stopCode; const lastStop = vr.aggregatedStops[vr.aggregatedStops.length - 1]?.stopCode; return firstStop === originCode && lastStop === destCode; },
                (vr) => { const lane = vr?.simpleFacilityLane || vr?.facilityLaneString || ''; return lane.includes(originCode) && lane.includes(destCode); }
            ];
            for (const strategy of strategies) { const match = fmcRecords.find(strategy); if (match) { Logger.debug('FMC API: Found matching VR', { vrId: match.vehicleRunId, lane: match.simpleFacilityLane }); return match; } }
            if (fmcRecords.length > 0) { Logger.warn('FMC API: No exact lane match found, using first VR'); return fmcRecords[0]; }
            return null;
        },

        extractTimestamps(vrRecord) {
            if (!vrRecord?.aggregatedStops || vrRecord.aggregatedStops.length === 0) { Logger.warn('FMC API: No stops available for timestamp extraction'); return null; }
            const stops = vrRecord.aggregatedStops;
            const originStop = stops.find(s => !s?.lastStop) || stops[0];
            const destStop = stops.find(s => s?.lastStop) || stops[stops.length - 1];
            if (!originStop || !destStop) { Logger.warn('FMC API: Could not determine origin/destination stops'); return null; }

            const getActionTimestamps = (stop) => {
                if (!stop?.actions?.length) {
                    return { plannedYardArrival: null, plannedYardDeparture: null, actualYardArrival: null, actualYardDeparture: null, arrivalSource: null, departureSource: null, arrivalDelayReason: null, departureDelayReason: null, actualYardArrivalSourceVrId: null, actualYardDepartureSourceVrId: null };
                }
                const action = stop.actions[0];
                return {
                    plannedYardArrival: action?.plannedYardArrivalTime, plannedYardDeparture: action?.plannedYardDepartureTime,
                    actualYardArrival: action?.actualYardArrivalTime, actualYardDeparture: action?.actualYardDepartureTime,
                    arrivalSource: action?.actualYardArrivalTimeSource, departureSource: action?.actualYardDepartureTimeSource,
                    arrivalDelayReason: action?.arrivalDelayReason, departureDelayReason: action?.departureDelayReason,
                    actualYardArrivalSourceVrId: null, actualYardDepartureSourceVrId: null
                };
            };

            return {
                origin: { stopCode: originStop?.stopCode || 'Unknown', displayName: originStop?.displayName || originStop?.stopCode || 'Unknown', timezone: originStop?.timezone || 'America/Los_Angeles', status: originStop?.status, actionType: originStop?.actions?.[0]?.type || 'PICKUP', timestamps: getActionTimestamps(originStop) },
                destination: { stopCode: destStop?.stopCode || 'Unknown', displayName: destStop?.displayName || destStop?.stopCode || 'Unknown', timezone: destStop?.timezone || 'America/Los_Angeles', status: destStop?.status, actionType: destStop?.actions?.[0]?.type || 'DROPOFF', timestamps: getActionTimestamps(destStop) },
                vrMetadata: { vehicleRunId: vrRecord?.vehicleRunId, tourId: vrRecord?.tourId, executionStatus: vrRecord?.executionStatus, hasDelays: vrRecord?.hasDelays, carrierName: vrRecord?.carrierName, totalDistance: vrRecord?.totalDistance }
            };
        },

        fillMissingTimestampsFromTour(timestamps, allRecords, currentVrId) {
            if (!timestamps || !allRecords || allRecords.length === 0) return timestamps;

            try { timestamps = JSON.parse(JSON.stringify(timestamps)); }
            catch (e) { Logger.warn('fillMissingTimestampsFromTour: Failed to clone timestamps', e.message); return timestamps; }

            const findTimestampFromOtherVRs = (stopCode, missingFields) => {
                const borrowedData = {};
                for (const vr of allRecords) {
                    if (vr?.vehicleRunId === currentVrId) continue;
                    if (!vr?.aggregatedStops) continue;
                    const matchingStop = vr.aggregatedStops.find(stop => stop?.stopCode === stopCode);
                    if (matchingStop?.actions?.length > 0) {
                        const action = matchingStop.actions[0];
                        for (const field of missingFields) { if (!borrowedData[field] && action?.[field]) { borrowedData[field] = { value: action[field], sourceVrId: vr.vehicleRunId }; } }
                        if (Object.keys(borrowedData).length === missingFields.length) break;
                    }
                }
                return borrowedData;
            };

            if (timestamps.origin?.stopCode) {
                const originMissing = [];
                if (!timestamps.origin.timestamps?.actualYardArrival) originMissing.push('actualYardArrivalTime');
                if (!timestamps.origin.timestamps?.actualYardDeparture) originMissing.push('actualYardDepartureTime');
                if (originMissing.length > 0) {
                    const borrowedOrigin = findTimestampFromOtherVRs(timestamps.origin.stopCode, originMissing);
                    if (borrowedOrigin.actualYardArrivalTime && timestamps.origin.timestamps) { timestamps.origin.timestamps.actualYardArrival = borrowedOrigin.actualYardArrivalTime.value; timestamps.origin.timestamps.actualYardArrivalSourceVrId = borrowedOrigin.actualYardArrivalTime.sourceVrId; }
                    if (borrowedOrigin.actualYardDepartureTime && timestamps.origin.timestamps) { timestamps.origin.timestamps.actualYardDeparture = borrowedOrigin.actualYardDepartureTime.value; timestamps.origin.timestamps.actualYardDepartureSourceVrId = borrowedOrigin.actualYardDepartureTime.sourceVrId; }
                }
            }

            if (timestamps.destination?.stopCode) {
                const destMissing = [];
                if (!timestamps.destination.timestamps?.actualYardArrival) destMissing.push('actualYardArrivalTime');
                if (!timestamps.destination.timestamps?.actualYardDeparture) destMissing.push('actualYardDepartureTime');
                if (destMissing.length > 0) {
                    const borrowedDest = findTimestampFromOtherVRs(timestamps.destination.stopCode, destMissing);
                    if (borrowedDest.actualYardArrivalTime && timestamps.destination.timestamps) { timestamps.destination.timestamps.actualYardArrival = borrowedDest.actualYardArrivalTime.value; timestamps.destination.timestamps.actualYardArrivalSourceVrId = borrowedDest.actualYardArrivalTime.sourceVrId; }
                    if (borrowedDest.actualYardDepartureTime && timestamps.destination.timestamps) { timestamps.destination.timestamps.actualYardDeparture = borrowedDest.actualYardDepartureTime.value; timestamps.destination.timestamps.actualYardDepartureSourceVrId = borrowedDest.actualYardDepartureTime.sourceVrId; }
                }
            }

            return timestamps;
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 23: SMC API SERVICE (Fix #2: uses _authenticatedRequest)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const SMCApiService = {
        async fetchOrderFull(orderId) {
            Logger.debug(`SMC API: Fetching full order ${orderId}`);
            return await HttpClient.request({
                method: 'GET',
                url: `${CONFIG.URLS.SMC_BASE}/shipper/order/${encodeURIComponent(orderId)}`,
                headers: { 'Accept': 'application/json' },
                timeout: CONFIG.API.REQUEST_TIMEOUT
            }, 'Order Full', circuitBreakers.smc);
        },

        /**
         * Fix #2: Shared helper for authenticated POST requests with retry/circuit-breaker.
         * Handles token injection, 403→clear, and delegates to HttpClient.
         */
        async _authenticatedRequest(url, payload, context) {
            const token = TokenManager.getToken();
            if (!token) throw ErrorHandler.create(ErrorType.AUTH, Messages.ERRORS.TOKEN_MISSING, { action: context });
            return await HttpClient.request({
                method: 'POST',
                url,
                headers: {
                    'Accept': 'application/json, text/plain, */*',
                    'Content-Type': 'application/json',
                    'x-csrf-token': token
                },
                body: JSON.stringify(payload),
                timeout: CONFIG.API.REQUEST_TIMEOUT
            }, context, circuitBreakers.smc);
        },

        async updateOrder(orderData, newPricing) {
            const orderId = orderData?.orderId?.id;
            if (!orderId) throw ErrorHandler.create(ErrorType.VALIDATION, 'Order ID is required for update', { orderData });
            Logger.debug(`SMC API: Updating order ${orderId}`);
            const payload = {
                orderId: orderData.orderId, orderDetails: { ...orderData.orderDetails, shipperPricing: { ...orderData.orderDetails?.shipperPricing, pricing: newPricing } },
                vrId: orderData.vrId, tpId: orderData.tpId, auditDetails: orderData.auditDetails, orderStatus: orderData.orderStatus,
                invoiceStatus: orderData.invoiceStatus, executionStatus: orderData.executionStatus, executionSourceType: orderData.executionSourceType,
                orderCreationSource: orderData.orderCreationSource, invoiceNumbers: orderData.invoiceNumbers || [], invoiceDetails: orderData.invoiceDetails,
                requiresManualPlanChanges: orderData.requiresManualPlanChanges, orderAction: orderData.orderAction,
                orderExecutionItineraryVersion: orderData.orderExecutionItineraryVersion, shipmentList: orderData.shipmentList,
                tenderDecision: orderData.tenderDecision, businessIdentifier: orderData.businessIdentifier,
                validationFailureReasonCodes: orderData.validationFailureReasonCodes || [], tasks: []
            };
            return this._authenticatedRequest(
                `${CONFIG.URLS.SMC_BASE}/shipper/order/update`,
                payload,
                `Order Update (${orderId})`
            );
        },

        async addPricingLine(orderId, orderVersion, pricingConfig, chargeAmount) {
            Logger.debug(`SMC API: Adding pricing line to ${orderId}`, { pricingCode: pricingConfig.pricingCode, amount: chargeAmount });
            const payload = {
                pricingComponents: [{ price: { value: chargeAmount, unit: 'USD' }, reasonCode: 'ORIGINAL', pricingCode: pricingConfig.pricingCode, pricingId: null, pricingComponentId: null, audit: null, invoiceNumber: null, chargeStatus: null, invoiceNote: null, authorizationNumber: null, chargeDocuments: [], taxComponents: [], itemized: [], description: pricingConfig.description, type: 'ACCESSORIAL' }],
                orderId: { id: orderId, version: orderVersion }
            };
            return this._authenticatedRequest(
                `${CONFIG.URLS.SMC_BASE}/shipper/order/add-shipper-pricing`,
                payload,
                `Add Pricing (${orderId})`
            );
        },

        async addComment(orderId, comment) {
            if (!comment) { Logger.warn('addComment: Missing comment', { orderId }); return false; }
            const token = TokenManager.getToken();
            if (!token) { Logger.warn('addComment: Missing token', { orderId }); return false; }
            Logger.debug(`SMC API: Adding comment to ${orderId}`);
            try {
                await this._authenticatedRequest(
                    `${CONFIG.URLS.SMC_BASE}/shipper/order/comments/add`,
                    { orderId, comment },
                    `Add Comment (${orderId})`
                );
                return true;
            } catch (error) {
                Logger.warn('addComment failed', { orderId, error: error.message });
                return false;
            }
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 24: DETENTION ANALYZER (with lane mismatch support)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const DetentionAnalyzer = {
        analyze(stop, fmcTimestamps, sowConfig, orderStatus, holdInfo, stopIndex, shipperName, laneMismatchInfo) {
            const stopTypeInfo = DataHelpers.formatStopType(stop?.stopActionType);
            const loadTypeInfo = DataHelpers.formatLoadType(stop?.loadingType);
            const isPickup = stopTypeInfo.isPickup;

            const hasHold = isPickup ? holdInfo.shipper : holdInfo.receiver;
            const holdCode = isPickup ? holdInfo.shipperCode : holdInfo.receiverCode;
            const detentionExists = isPickup ? holdInfo.shipperExists : holdInfo.receiverExists;
            const existingAmount = isPickup ? holdInfo.shipperAmount : holdInfo.receiverAmount;

            const result = {
                type: null, charge: 0, breakdown: '', hitMax: false, action: ActionType.NO_ACTION, actionText: '', comment: '',
                hasHold, holdCode, detentionExists, existingAmount, isPickup, fmcTimestamps,
                requiresApproval: sowConfig?.requiresApproval || false, autoChargeAllowed: sowConfig?.autoChargeAllowed || false,
                authNumberRequired: sowConfig?.authNumberRequired || false, sowConfig,
                processed: false, processedAction: null, processedAmount: null, processError: null,
                hasBorrowedTimestamps: false, approvalNote: null,
                calculatedAmount: null, mismatchAmount: null, mismatchType: null, orderStatusGroup: null, noChargeReason: null,
                laneMismatch: false, laneValidation: null
            };

            const statusInfo = DataHelpers.formatStatusDisplay(orderStatus);
            result.orderStatusGroup = statusInfo.group;

            // Lane mismatch — applies to ALL stops (whole order)
            if (laneMismatchInfo && !laneMismatchInfo.isMatch) {
                result.laneMismatch = true;
                result.laneValidation = laneMismatchInfo;
                result.type = ResultType.LANE_MISMATCH;
                result.breakdown = `Lane mismatch: SMC (${laneMismatchInfo.smcOrigin} → ${laneMismatchInfo.smcDest}) ≠ FMC (${laneMismatchInfo.fmcOrigin} → ${laneMismatchInfo.fmcDest})`;
                if (hasHold) {
                    result.action = ActionType.RELEASE_HOLD;
                    result.actionText = 'Release hold (Lane Mismatch)';
                    result.comment = Messages.COMMENTS.RELEASE_HOLD; // Timestamp comment (Comment A)
                } else {
                    result.action = ActionType.NO_ACTION;
                    result.actionText = 'No charge (Lane Mismatch)';
                }
                return result;
            }

            if (statusInfo.group === 'cancelled') return this._setResult(result, ResultType.ORDER_CANCELLED, 'Order was cancelled/rejected', ActionType.NO_ACTION, 'Order cancelled');

            if (statusInfo.group === 'invoiced' || statusInfo.group === 'paid') {
                if (ModeManager.isReviewMode() && existingAmount > 0) {
                    Logger.debug(`Review Mode: Bypassing invoiced/paid early return for mismatch check (existing: $${existingAmount})`);
                } else {
                    return this._setResult(result, ResultType.ORDER_INVOICED, 'Order already invoiced', ActionType.NO_ACTION, 'Already invoiced');
                }
            }

            if (existingAmount > 0) {
                if (ModeManager.isReviewMode()) {
                    Logger.debug(`Review Mode: Continuing past charge-exists for mismatch check (existing: $${existingAmount})`);
                } else {
                    result.type = ResultType.CHARGE_EXISTS;
                    result.charge = existingAmount;
                    result.breakdown = `Detention charge already exists: $${existingAmount.toFixed(2)}\nNo modifications will be made.`;
                    result.action = ActionType.NO_ACTION;
                    result.actionText = `Charge already added ($${existingAmount.toFixed(2)})`;
                    return result;
                }
            }

            if (!fmcTimestamps?.timestamps) {
                if (ModeManager.isReviewMode() && existingAmount > 0) {
                    return this._buildMismatchResult(result, existingAmount, null, 'FMC data unavailable');
                }
                return this._setResult(result, ResultType.FMC_DATA_UNAVAILABLE, 'FMC timestamp data not available for this stop', ActionType.PENDING, 'FMC data unavailable');
            }

            const ts = fmcTimestamps.timestamps;

            if (!ts.actualYardArrival) {
                if (ModeManager.isReviewMode() && existingAmount > 0) {
                    return this._buildMismatchResult(result, existingAmount, null, 'Awaiting arrival');
                }
                return this._setResult(result, ResultType.MISSING_ARRIVAL, 'Driver has not arrived yet (FMC)', ActionType.PENDING, Messages.INFO.AWAITING_ARRIVAL);
            }

            if (!ts.actualYardDeparture) {
                if (ModeManager.isReviewMode() && existingAmount > 0) {
                    return this._buildMismatchResult(result, existingAmount, null, 'Awaiting departure');
                }
                return this._setResult(result, ResultType.MISSING_DEPARTURE, 'Driver has arrived but not departed yet (FMC)', ActionType.PENDING, Messages.INFO.AWAITING_DEPARTURE);
            }

            const stopKey = stopTypeInfo.key;
            const loadKey = loadTypeInfo.key;
            const rules = sowConfig?.rules?.[stopKey]?.[loadKey];
            if (!rules) return this._setResult(result, ResultType.UNKNOWN_ERROR, `No rules found for ${stopKey}/${loadKey}`, ActionType.ERROR, 'SOW rule missing');

            if (!rules.eligible) {
                if (ModeManager.isReviewMode() && existingAmount > 0) {
                    return this._buildMismatchResult(result, existingAmount, 0, 'Not eligible per SOW');
                }
                result.type = ResultType.NO_DETENTION_DROP_HOOK;
                result.breakdown = `${loadTypeInfo.display} - Not eligible for detention per SOW`;
                if (hasHold) { result.action = ActionType.RELEASE_HOLD; result.actionText = 'Release hold (Not eligible)'; result.comment = Messages.COMMENTS.RELEASE_HOLD; }
                else { result.action = ActionType.NO_ACTION; result.actionText = 'No detention (Not eligible)'; }
                return result;
            }

            const plannedArrival = ts.plannedYardArrival;
            const actualArrival = ts.actualYardArrival;
            const arrivalDiff = DataHelpers.calculateTimeDiff(plannedArrival, actualArrival);

            if (arrivalDiff.minutes !== null && arrivalDiff.minutes > CONFIG.TIMING.LATE_MINUTES) {
                if (ModeManager.isReviewMode() && existingAmount > 0) {
                    return this._buildMismatchResult(result, existingAmount, 0, 'Driver late');
                }
                result.type = ResultType.DRIVER_LATE;
                result.breakdown = `Driver arrived ${Helpers.formatDuration(arrivalDiff.minutes)} late (FMC)\nNot eligible for detention`;
                if (hasHold) { result.action = ActionType.RELEASE_HOLD; result.actionText = 'Release hold (Late)'; result.comment = Messages.COMMENTS.RELEASE_HOLD; }
                else { result.action = ActionType.NO_ACTION; result.actionText = Messages.INFO.DRIVER_LATE; }
                return result;
            }

            const plannedDeparture = ts.plannedYardDeparture;
            const actualDeparture = ts.actualYardDeparture;
            const departureDiff = DataHelpers.calculateTimeDiff(plannedDeparture, actualDeparture);
            const delayMinutes = (departureDiff.minutes !== null && departureDiff.minutes > 0) ? departureDiff.minutes : 0;

            const chargeResult = this._calculateCharge(result, sowConfig, rules, delayMinutes, hasHold);

            if (ModeManager.isReviewMode() && existingAmount > 0) {
                const calculatedCharge = chargeResult.charge;
                if (chargeResult.type === ResultType.WITHIN_FREE_TIME || chargeResult.type === ResultType.NO_HOLD_NO_CHARGE || chargeResult.type === ResultType.BELOW_MINIMUM_THRESHOLD) {
                    return this._buildMismatchResult(chargeResult, existingAmount, 0, chargeResult.type === ResultType.WITHIN_FREE_TIME ? 'Within free time' : chargeResult.type === ResultType.BELOW_MINIMUM_THRESHOLD ? 'Below minimum' : 'No charge');
                }
                return this._buildMismatchResult(chargeResult, existingAmount, calculatedCharge, null);
            }

            return chargeResult;
        },

        _setResult(result, type, breakdown, action, actionText) {
            result.type = type; result.breakdown = breakdown; result.action = action; result.actionText = actionText;
            return result;
        },

        _buildMismatchResult(result, existingAmount, calculatedAmount, noChargeReason) {
            result.existingAmount = existingAmount;
            result.calculatedAmount = calculatedAmount;
            result.noChargeReason = noChargeReason;
            result.action = ActionType.NO_ACTION;

            if (calculatedAmount === null) {
                result.type = ResultType.CHARGE_EXISTS;
                result.mismatchType = null;
                result.actionText = `$${existingAmount.toFixed(2)} (Cannot verify - ${noChargeReason})`;
                return result;
            }

            if (calculatedAmount === 0) {
                result.type = ResultType.CHARGE_SHOULD_NOT_EXIST;
                result.mismatchType = 'SHOULD_NOT_EXIST';
                result.mismatchAmount = existingAmount;
                result.charge = 0;
                result.actionText = `🚨 Should not be charged (${noChargeReason})`;
                return result;
            }

            if (existingAmount === calculatedAmount) {
                result.type = ResultType.CHARGE_CORRECT;
                result.mismatchType = 'CORRECT';
                result.mismatchAmount = 0;
                result.actionText = `✅ Charge correct ($${existingAmount.toFixed(2)})`;
                return result;
            }

            if (existingAmount < calculatedAmount) {
                result.type = ResultType.CHARGE_MISMATCH_UNDER;
                result.mismatchType = 'UNDER';
                result.mismatchAmount = calculatedAmount - existingAmount;
                result.actionText = `⚠️ Undercharged by $${result.mismatchAmount.toFixed(2)}`;
                return result;
            }

            result.type = ResultType.CHARGE_MISMATCH_OVER;
            result.mismatchType = 'OVER';
            result.mismatchAmount = existingAmount - calculatedAmount;
            result.actionText = `⚠️ Overcharged by $${result.mismatchAmount.toFixed(2)}`;
            return result;
        },

        _calculateCharge(result, sowConfig, rules, delayMinutes, hasHold) {
            const freeTime = rules.freeTime || 0;
            let chargeableMinutes = delayMinutes - freeTime;

            if (chargeableMinutes <= 0) {
                result.type = ResultType.WITHIN_FREE_TIME;
                result.breakdown = `Delay: ${delayMinutes} min (FMC)\nFree Time: ${freeTime} min\nNo charge`;
                if (hasHold) { result.action = ActionType.RELEASE_HOLD; result.actionText = 'Release hold ($0)'; result.comment = Messages.COMMENTS.RELEASE_HOLD; }
                else { result.type = ResultType.NO_HOLD_NO_CHARGE; result.action = ActionType.NO_ACTION; result.actionText = Messages.INFO.NO_ACTION_NEEDED; }
                return result;
            }

            if (sowConfig.roundDownMaxMinutes && sowConfig.roundDownMaxMinutes > 0) {
                if (chargeableMinutes < sowConfig.roundDownMaxMinutes) {
                    result.type = ResultType.BELOW_MINIMUM_THRESHOLD;
                    result.breakdown = [`Delay: ${delayMinutes} min (FMC)`, `Free Time: -${freeTime} min`, `Chargeable: ${chargeableMinutes} min`, `Minimum Threshold: ${sowConfig.roundDownMaxMinutes} min`, `Below threshold - No charge`].join('\n');
                    if (hasHold) { result.action = ActionType.RELEASE_HOLD; result.actionText = 'Release hold (Below min)'; result.comment = Messages.COMMENTS.RELEASE_HOLD; }
                    else { result.action = ActionType.NO_ACTION; result.actionText = Messages.INFO.BELOW_MINIMUM; }
                    return result;
                }
            }

            const originalChargeableMinutes = chargeableMinutes;
            let wasRounded = false;

            if (sowConfig.billingIncrement && sowConfig.billingIncrement > 0 && sowConfig.roundingRule && sowConfig.roundingRule !== '') {
                chargeableMinutes = this._applyBillingIncrement(chargeableMinutes, sowConfig.billingIncrement, sowConfig.roundingRule);
                wasRounded = (chargeableMinutes !== originalChargeableMinutes);
            }

            let charge;
            if (sowConfig.rateUnit === 'MINUTE') charge = chargeableMinutes * sowConfig.rate;
            else charge = (chargeableMinutes / 60) * sowConfig.rate;
            charge = Math.round(charge * 100) / 100;

            const hitMax = charge >= sowConfig.maxCharge;
            charge = Math.min(charge, sowConfig.maxCharge);
            result.charge = charge; result.hitMax = hitMax;

            const breakdownLines = [`Delay: ${delayMinutes} min (FMC)`, `Free Time: -${freeTime} min`, `Chargeable: ${originalChargeableMinutes} min`];
            if (wasRounded) { breakdownLines.push(`Billing Increment: ${sowConfig.billingIncrement} min (${sowConfig.roundingRule})`); breakdownLines.push(`Rounded to: ${chargeableMinutes} min`); }
            if (sowConfig.rateUnit === 'MINUTE') breakdownLines.push(`Rate: $${sowConfig.rate}/min`);
            else breakdownLines.push(`Rate: $${sowConfig.rate}/hr`);
            breakdownLines.push(`Charge: $${charge.toFixed(2)}${hitMax ? ' (MAX)' : ''}`);
            result.breakdown = breakdownLines.join('\n');

            return this._determineAction(result, charge, hasHold, hitMax, sowConfig);
        },

        _applyBillingIncrement(minutes, increment, roundingRule) {
            if (!increment || increment <= 0) return minutes;
            const remainder = minutes % increment;
            if (remainder === 0) return minutes;
            const normalizedRule = String(roundingRule || '').toUpperCase();
            switch (normalizedRule) {
                case 'UP': return minutes + (increment - remainder);
                case 'DOWN': return minutes - remainder;
                case 'NEAREST': { const halfIncrement = increment / 2; if (remainder >= halfIncrement) return minutes + (increment - remainder); else return minutes - remainder; }
                default: Logger.warn(`Unknown rounding rule: ${roundingRule}, defaulting to UP`); return minutes + (increment - remainder);
            }
        },

        _determineAction(result, charge, hasHold, hitMax, sowConfig) {
            const autoChargeAllowed = sowConfig?.autoChargeAllowed === true;
            const requiresApproval = sowConfig?.requiresApproval === true;

            if (autoChargeAllowed) {
                if (requiresApproval) { result.type = ResultType.CHARGEABLE; result.action = ActionType.PENDING_APPROVAL; result.actionText = `Approval needed: $${charge.toFixed(2)}${hitMax ? ' (MAX)' : ''}`; return result; }
                result.type = hitMax ? ResultType.CHARGEABLE_MAX : ResultType.CHARGEABLE;
                if (hasHold) { result.action = ActionType.ADD_CHARGE_UPDATE; result.actionText = `Update to $${charge.toFixed(2)}${hitMax ? ' (MAX)' : ''}`; result.comment = Messages.COMMENTS.ADD_CHARGE; }
                else { result.action = ActionType.ADD_CHARGE_CREATE; result.actionText = `Recover $${charge.toFixed(2)}${hitMax ? ' (MAX)' : ''}`; result.comment = Messages.COMMENTS.ADD_CHARGE; }
                return result;
            }

            result.type = hitMax ? ResultType.CHARGEABLE_MAX : ResultType.CHARGEABLE;
            if (requiresApproval) { result.action = ActionType.PENDING_APPROVAL; result.actionText = `Approval needed: $${charge.toFixed(2)}${hitMax ? ' (MAX)' : ''}`; return result; }
            if (hasHold) { result.action = ActionType.ANALYSIS_ONLY; result.actionText = `Analysis only: $${charge.toFixed(2)}${hitMax ? ' (MAX)' : ''} (Hold exists)`; }
            else { result.action = ActionType.CREATE_HOLD_ONLY; result.actionText = `Create $0 hold: $${charge.toFixed(2)}${hitMax ? ' (MAX)' : ''} chargeable`; }
            return result;
        },

        previewCharge(params) {
            const { delayMinutes, freeTime, rate, rateUnit, maxCharge, billingIncrement, roundingRule, roundDownMaxMinutes } = params;
            let chargeableMinutes = delayMinutes - (freeTime || 0);
            if (chargeableMinutes <= 0) return { charge: 0, chargeableMinutes: 0, withinFreeTime: true };
            if (roundDownMaxMinutes && chargeableMinutes < roundDownMaxMinutes) return { charge: 0, chargeableMinutes, belowMinimum: true };
            const originalMinutes = chargeableMinutes;
            if (billingIncrement && billingIncrement > 0 && roundingRule) chargeableMinutes = this._applyBillingIncrement(chargeableMinutes, billingIncrement, roundingRule);
            let charge;
            if (rateUnit === 'MINUTE') charge = chargeableMinutes * rate;
            else charge = (chargeableMinutes / 60) * rate;
            charge = Math.round(charge * 100) / 100;
            const hitMax = charge >= maxCharge;
            charge = Math.min(charge, maxCharge);
            return { charge, chargeableMinutes, originalMinutes, wasRounded: chargeableMinutes !== originalMinutes, hitMax };
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 25: HTML GENERATOR (with lane mismatch display)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const HTMLGenerator = {
        getStatusIcon(status) {
            switch (status) { case ShipperStatus.ACTIVE: return '✅'; case ShipperStatus.INACTIVE: return '⏸️'; case ShipperStatus.VALIDATION_ERROR: return '❌'; default: return '❓'; }
        },

        getStatusClass(status) {
            switch (status) { case ShipperStatus.ACTIVE: return 'status-active'; case ShipperStatus.INACTIVE: return 'status-inactive'; case ShipperStatus.VALIDATION_ERROR: return 'status-error'; default: return ''; }
        },

        renderShipperCard(shipper, isExpanded) {
            const statusIcon = HTMLGenerator.getStatusIcon(shipper.status);
            const statusClass = HTMLGenerator.getStatusClass(shipper.status);
            const data = shipper.rawData;
            const escapedName = Helpers.escapeHtml(shipper.shipperName);
            const safeId = escapedName.replace(/[^a-zA-Z0-9-]/g, '-');

            return `
                <div class="d-dart-shipper-card-settings ${statusClass}" data-shipper="${escapedName}">
                    <div class="d-dart-shipper-header-settings" data-toggle-shipper="${escapedName}">
                        <div class="d-dart-shipper-info">
                            <span class="d-dart-shipper-status-icon">${statusIcon}</span>
                            <span class="d-dart-shipper-name-settings">${escapedName}</span>
                        </div>
                        <div class="d-dart-shipper-summary">
                            ${shipper.status !== ShipperStatus.VALIDATION_ERROR ? `
                                <span class="d-dart-shipper-rate">💰 ${Helpers.formatValueOrDash(data.Rate ? `$${data.Rate}/${String(data.RateUnit || 'hr').toLowerCase()}` : null)}</span>
                                <span class="d-dart-shipper-max">🔝 ${Helpers.formatValueOrDash(data.MaxCharge ? `$${data.MaxCharge}` : null)}</span>
                            ` : `<span class="d-dart-validation-error-badge">⚠️ Validation Error</span>`}
                            <button class="d-dart-expand-btn" aria-label="${isExpanded ? 'Collapse' : 'Expand'} details">${isExpanded ? '▲' : '▼'}</button>
                        </div>
                    </div>
                    <div class="d-dart-shipper-details ${isExpanded ? 'expanded' : ''}" id="d-dart-details-${safeId}">
                        ${shipper.status === ShipperStatus.VALIDATION_ERROR ? `
                            <div class="d-dart-validation-errors"><div class="d-dart-error-title">❌ Validation Errors:</div><ul class="d-dart-error-list">${shipper.validationErrors.map(err => `<li>${Helpers.escapeHtml(err)}</li>`).join('')}</ul></div>
                        ` : ''}
                        <div class="d-dart-details-grid">
                            <div class="d-dart-details-section"><div class="d-dart-details-title">📌 BASIC INFO</div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Rate:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.Rate)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Unit:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.RateUnit)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Max:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.MaxCharge ? `$${data.MaxCharge}` : null)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Active:</span><span class="d-dart-details-value ${parseBoolean(data.IsActive, true) ? 'yes' : 'no'}">${Helpers.formatBoolean(data.IsActive)}</span></div>
                            </div>
                            <div class="d-dart-details-section"><div class="d-dart-details-title">📍 PU LIVE</div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Eligible:</span><span class="d-dart-details-value">${Helpers.formatBoolean(data.PU_Live_Eligible)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">FreeTime:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.PU_Live_FreeTime != null ? `${data.PU_Live_FreeTime} min` : null)}</span></div>
                            </div>
                            <div class="d-dart-details-section"><div class="d-dart-details-title">📍 PU DROP HOOK</div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Eligible:</span><span class="d-dart-details-value">${Helpers.formatBoolean(data.PU_DropHook_Eligible)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">FreeTime:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.PU_DropHook_FreeTime != null ? `${data.PU_DropHook_FreeTime} min` : null)}</span></div>
                            </div>
                            <div class="d-dart-details-section"><div class="d-dart-details-title">📍 DO LIVE</div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Eligible:</span><span class="d-dart-details-value">${Helpers.formatBoolean(data.DO_Live_Eligible)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">FreeTime:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.DO_Live_FreeTime != null ? `${data.DO_Live_FreeTime} min` : null)}</span></div>
                            </div>
                            <div class="d-dart-details-section"><div class="d-dart-details-title">📍 DO DROP HOOK</div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Eligible:</span><span class="d-dart-details-value">${Helpers.formatBoolean(data.DO_DropHook_Eligible)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">FreeTime:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.DO_DropHook_FreeTime != null ? `${data.DO_DropHook_FreeTime} min` : null)}</span></div>
                            </div>
                            <div class="d-dart-details-section"><div class="d-dart-details-title">⚙️ OPTIONAL</div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Billing:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.BillingIncrement ? `${data.BillingIncrement} min` : null)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Rounding:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.RoundingRule)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">MinThreshold:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.RoundDownMaxMinutes ? `${data.RoundDownMaxMinutes} min` : null)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Approval:</span><span class="d-dart-details-value">${Helpers.formatBoolean(data.RequiresApproval)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">AutoCharge:</span><span class="d-dart-details-value">${Helpers.formatBoolean(data.AutoChargeAllowed)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">AuthReq:</span><span class="d-dart-details-value">${Helpers.formatBoolean(data.AuthNumberRequired)}</span></div>
                                <div class="d-dart-details-row"><span class="d-dart-details-label">Notes:</span><span class="d-dart-details-value">${Helpers.formatValueOrDash(data.Notes)}</span></div>
                            </div>
                        </div>
                    </div>
                </div>
            `;
        },

        renderShippersList() {
            const expandedShippers = AppState.get('expandedShippers');
            const shippers = SOWConfigManager.getAllShippersData();
            if (shippers.length === 0) return '<div class="d-dart-no-results">No shippers configured</div>';
            return shippers.map(shipper => HTMLGenerator.renderShipperCard(shipper, expandedShippers.has(shipper.shipperName))).join('');
        },

        shippersLoadingSkeleton() {
            return `
                <div class="d-dart-shippers-loading">
                    <div class="d-dart-loading-skeleton"></div>
                    <div class="d-dart-loading-skeleton" style="animation-delay: 0.15s"></div>
                    <div class="d-dart-loading-skeleton" style="animation-delay: 0.3s"></div>
                    <div class="d-dart-loading-text">🔄 Fetching from SharePoint...</div>
                </div>
            `;
        },

        settingsPanelInline() {
            const stats = SOWConfigManager.getStatistics();
            const lastRefresh = SOWConfigManager.getLastRefreshTime();
            const lastRefreshText = lastRefresh ? Helpers.formatRelativeTime(lastRefresh) : 'Never';

            return `
                <div class="d-dart-settings-inline" id="d-dart-settings-inline">
                    <div class="d-dart-settings-inline-header">
                        <button class="d-dart-back-btn" id="d-dart-settings-back" aria-label="Back to main view">← Back</button>
                        <span class="d-dart-settings-inline-title">⚙️ SOW SETTINGS</span>
                    </div>
                    <div class="d-dart-settings-refresh-section">
                        <button class="d-dart-refresh-sow-btn" id="d-dart-refresh-sow">🔄 Refresh SOW Configuration</button>
                        <div class="d-dart-refresh-elapsed" id="d-dart-refresh-elapsed">Last refresh: ${Helpers.escapeHtml(lastRefreshText)}</div>
                    </div>
                    <div class="d-dart-settings-section">
                        <div class="d-dart-settings-section-title" id="d-dart-shippers-count">📋 SHIPPERS (${stats.total})</div>
                        <div class="d-dart-shippers-list" id="d-dart-shippers-list">${HTMLGenerator.renderShippersList()}</div>
                    </div>
                </div>
            `;
        },

        getBannerValueClass(analysis) {
            if (!analysis) return 'no-action';
            if (analysis.laneMismatch) return 'lane-mismatch';
            if (analysis.mismatchType) {
                switch (analysis.mismatchType) {
                    case 'UNDER': return 'mismatch-under';
                    case 'OVER': return 'mismatch-over';
                    case 'SHOULD_NOT_EXIST': return 'mismatch-should-not-exist';
                    case 'CORRECT': return 'charge-correct';
                }
            }
            if (analysis.processed) {
                switch (analysis.processedAction) {
                    case 'updated': case 'created': return 'charge-added';
                    case 'released': return 'hold-released';
                    case 'released_lane_mismatch': return 'lane-mismatch';
                    case 'hold_created': return 'hold-created';
                    case 'analysis_only': return 'analysis-only';
                    default: return 'no-action';
                }
            }
            switch (analysis.type) {
                case ResultType.LANE_MISMATCH: return 'lane-mismatch';
                case ResultType.CHARGEABLE: case ResultType.CHARGEABLE_MAX:
                    if (analysis.action === ActionType.ANALYSIS_ONLY) return 'analysis-only';
                    if (analysis.action === ActionType.CREATE_HOLD_ONLY) return 'hold-created';
                    if (analysis.action === ActionType.PENDING_APPROVAL) return 'chargeable';
                    return 'chargeable';
                case ResultType.CHARGE_EXISTS: return 'charge-exists';
                case ResultType.MISSING_ARRIVAL: case ResultType.MISSING_DEPARTURE: case ResultType.FMC_DATA_UNAVAILABLE: return 'pending';
                case ResultType.DRIVER_LATE: case ResultType.ORDER_CANCELLED: case ResultType.BELOW_MINIMUM_THRESHOLD: return 'no-charge';
                default: return 'no-action';
            }
        },

        laneMismatchBanner(laneValidation) {
            if (!laneValidation || !laneValidation.validated) return '';
            if (laneValidation.isMatch) return '';
            return `
                <div class="d-dart-lane-mismatch-banner">
                    <div class="d-dart-lane-mismatch-icon">🚫</div>
                    <div class="d-dart-lane-mismatch-title">LANE MISMATCH DETECTED</div>
                    <div class="d-dart-lane-mismatch-detail">
                        <div class="d-dart-lane-compare-row">
                            <span class="d-dart-lane-compare-label">SMC Contracted:</span>
                            <span class="d-dart-lane-compare-value">${Helpers.escapeHtml(laneValidation.smcOrigin)} → ${Helpers.escapeHtml(laneValidation.smcDest)}</span>
                        </div>
                        <div class="d-dart-lane-compare-row">
                            <span class="d-dart-lane-compare-label">FMC Execution:</span>
                            <span class="d-dart-lane-compare-value">
                                <span class="${laneValidation.originMatch ? '' : 'd-dart-lane-mismatch-highlight'}">${Helpers.escapeHtml(laneValidation.fmcOrigin)}</span>
                                →
                                <span class="${laneValidation.destMatch ? '' : 'd-dart-lane-mismatch-highlight'}">${Helpers.escapeHtml(laneValidation.fmcDest)}</span>
                            </span>
                        </div>
                    </div>
                    <div class="d-dart-lane-mismatch-action">Detention not applicable — holds released with timestamp comment</div>
                </div>
            `;
        },

        detentionSummaryBanner(orderData) {
            const analysisResults = orderData?.analysisResults || [];
            const shipperAnalysis = analysisResults.find(a => a?.isPickup);
            const receiverAnalysis = analysisResults.find(a => !a?.isPickup);
            const shipperText = DataHelpers.getActionDisplayText(shipperAnalysis);
            const receiverText = DataHelpers.getActionDisplayText(receiverAnalysis);
            const toggleId = Helpers.generateId('detention-details');
            const shipperFmcData = shipperAnalysis?.fmcStopData || orderData?.fmcTimestamps?.origin;
            const receiverFmcData = receiverAnalysis?.fmcStopData || orderData?.fmcTimestamps?.destination;
            const shipperTimezone = shipperFmcData?.timezone || 'America/Los_Angeles';
            const receiverTimezone = receiverFmcData?.timezone || 'America/Los_Angeles';
            const shipperBreakdown = DataHelpers.getBreakdownDetails(shipperAnalysis, shipperFmcData, shipperTimezone);
            const receiverBreakdown = DataHelpers.getBreakdownDetails(receiverAnalysis, receiverFmcData, receiverTimezone);

            // Lane mismatch banner (shown above detention summary)
            const laneValidation = orderData?.laneValidation;
            const laneMismatchHtml = HTMLGenerator.laneMismatchBanner(laneValidation);

            const buildMismatchBox = (analysis) => {
                if (!analysis?.mismatchType || analysis.mismatchType === 'CORRECT') return '';
                const statusTag = analysis.orderStatusGroup === 'invoiced' ? ' [Invoiced]' : analysis.orderStatusGroup === 'paid' ? ' [Paid]' : '';
                const label = analysis.isPickup ? 'SHIPPER' : 'RECEIVER';
                if (analysis.mismatchType === 'SHOULD_NOT_EXIST') {
                    return `<div class="d-dart-mismatch-box should-not-exist"><div class="d-dart-mismatch-title">🚨 ${label}: Should Not Be Charged${statusTag}</div><div class="d-dart-mismatch-row"><span>Current in SMC:</span><span>$${analysis.existingAmount.toFixed(2)}</span></div><div class="d-dart-mismatch-row"><span>Correct Amount:</span><span>$0.00 (${Helpers.escapeHtml(analysis.noChargeReason || 'N/A')})</span></div><div class="d-dart-mismatch-row diff"><span>Difference:</span><span>-$${analysis.existingAmount.toFixed(2)} (Overcharged)</span></div></div>`;
                }
                const diffSign = analysis.mismatchType === 'UNDER' ? '+' : '-';
                const diffLabel = analysis.mismatchType === 'UNDER' ? 'Undercharged' : 'Overcharged';
                return `<div class="d-dart-mismatch-box ${analysis.mismatchType === 'UNDER' ? 'under' : 'over'}"><div class="d-dart-mismatch-title">⚠️ ${label}: Incorrect Charge Detected${statusTag}</div><div class="d-dart-mismatch-row"><span>Current in SMC:</span><span>$${analysis.existingAmount.toFixed(2)}</span></div><div class="d-dart-mismatch-row"><span>Correct Amount:</span><span>$${analysis.calculatedAmount.toFixed(2)}</span></div><div class="d-dart-mismatch-row diff"><span>Difference:</span><span>${diffSign}$${analysis.mismatchAmount.toFixed(2)} (${diffLabel})</span></div></div>`;
            };

            const shipperMismatchBox = buildMismatchBox(shipperAnalysis);
            const receiverMismatchBox = buildMismatchBox(receiverAnalysis);

            return `
                ${laneMismatchHtml}
                <div class="d-dart-detention-banner">
                    <div class="d-dart-banner-header">
                        <span class="d-dart-banner-title">📊 DETENTION SUMMARY</span>
                        <button class="d-dart-banner-toggle" data-toggle-target="${toggleId}" aria-expanded="false" aria-controls="${toggleId}">▼ Details</button>
                    </div>
                    ${shipperMismatchBox}${receiverMismatchBox}
                    <div class="d-dart-banner-summary">
                        <div class="d-dart-banner-row">
                            <span class="d-dart-banner-label">SHIPPER:</span>
                            <span class="d-dart-banner-value ${HTMLGenerator.getBannerValueClass(shipperAnalysis)}">${Helpers.escapeHtml(shipperText)}</span>
                        </div>
                        <div class="d-dart-banner-row">
                            <span class="d-dart-banner-label">RECEIVER:</span>
                            <span class="d-dart-banner-value ${HTMLGenerator.getBannerValueClass(receiverAnalysis)}">${Helpers.escapeHtml(receiverText)}</span>
                        </div>
                    </div>
                    <div class="d-dart-banner-details" id="${toggleId}" aria-hidden="true">
                        <div class="d-dart-breakdown-section"><div class="d-dart-breakdown-title">SHIPPER BREAKDOWN:</div><div class="d-dart-breakdown-content">${shipperBreakdown.map(line => `<div class="d-dart-breakdown-line">${Helpers.escapeHtml(line)}</div>`).join('')}</div></div>
                        <div class="d-dart-breakdown-section"><div class="d-dart-breakdown-title">RECEIVER BREAKDOWN:</div><div class="d-dart-breakdown-content">${receiverBreakdown.map(line => `<div class="d-dart-breakdown-line">${Helpers.escapeHtml(line)}</div>`).join('')}</div></div>
                    </div>
                </div>
            `;
        },

        sowDetails(sowConfig, toggleId) {
            if (!sowConfig) return '';
            const items = [
                { label: '💰 Rate:', value: sowConfig.displayInfo?.rate || '-' },
                { label: '🔝 Max:', value: sowConfig.displayInfo?.max || '-' },
                { label: '⏱️ Billing:', value: sowConfig.displayInfo?.billingIncrement || '-' },
                { label: '📐 Round:', value: sowConfig.roundingRule || '-' },
                { label: '🔄 PU Live:', value: `${sowConfig.rules?.PICKUP?.LIVE?.eligible ? '✓' : '✗'} ${sowConfig.rules?.PICKUP?.LIVE?.freeTime || 0}m` },
                { label: '🪝 PU D&H:', value: `${sowConfig.rules?.PICKUP?.DROP_HOOK?.eligible ? '✓' : '✗'} ${sowConfig.rules?.PICKUP?.DROP_HOOK?.freeTime || 0}m` },
                { label: '🔄 DO Live:', value: `${sowConfig.rules?.DROP_OFF?.LIVE?.eligible ? '✓' : '✗'} ${sowConfig.rules?.DROP_OFF?.LIVE?.freeTime || 0}m` },
                { label: '🪝 DO D&H:', value: `${sowConfig.rules?.DROP_OFF?.DROP_HOOK?.eligible ? '✓' : '✗'} ${sowConfig.rules?.DROP_OFF?.DROP_HOOK?.freeTime || 0}m` }
            ];
            if (sowConfig.requiresApproval) items.push({ label: '⚠️ Approval:', value: 'Required', class: 'warning' });
            if (sowConfig.autoChargeAllowed) items.push({ label: '⚡ Auto:', value: 'Enabled', class: 'success' });
            else items.push({ label: '⚡ Auto:', value: 'Disabled', class: 'disabled' });

            const itemsHtml = items.map(item => `<div class="d-dart-sow-item"><span class="d-dart-sow-item-label">${item.label}</span><span class="d-dart-sow-item-value ${item.class || ''}">${Helpers.escapeHtml(item.value)}</span></div>`).join('');
            return `<div class="d-dart-sow-details" id="${toggleId}" aria-hidden="true"><div class="d-dart-sow-flex">${itemsHtml}</div></div>`;
        },

        shipperCard(orderData) {
            const shipperName = orderData?.shipperName || 'Unknown Shipper';
            const truncatedName = Helpers.truncateText(shipperName, 25);
            const sowConfig = orderData?.sowConfig;
            const orderStatus = orderData?.viewData?.orderExecutionStatus || 'UNKNOWN';
            const statusInfo = DataHelpers.formatStatusDisplay(orderStatus);
            const orderId = orderData?.orderId || 'Unknown';
            const vrId = orderData?.viewData?.vehicleRunIds?.[0] || orderData?.viewData?.vrId || orderData?.smcExecutionData?.vehicleRunId || 'N/A';
            const tourId = orderData?.smcExecutionData?.tourId || 'N/A';
            const smcUrl = SecurityHelpers.buildSMCOrderUrl(orderId);
            const fmcVrUrl = SecurityHelpers.buildFMCSearchUrl(vrId);
            const fmcTourUrl = SecurityHelpers.buildFMCSearchUrl(tourId);
            const originCode = orderData?.smcExecutionData?.origin?.nodeCode || orderData?.fmcTimestamps?.origin?.stopCode || 'Unknown';
            const destCode = orderData?.smcExecutionData?.destination?.nodeCode || orderData?.fmcTimestamps?.destination?.stopCode || 'Unknown';
            const pricing = orderData?.viewData?.shipperPricing?.pricing || [];
            const holds = DataHelpers.findDetentionHolds(pricing);
            const sowToggleId = Helpers.generateId('sow');

            return `
                <div class="d-dart-shipper-card">
                    <div class="d-dart-header-row">
                        <div class="d-dart-shipper-name" title="${Helpers.escapeHtml(shipperName)}">🏢 ${Helpers.escapeHtml(truncatedName)}</div>
                        <div class="d-dart-header-badges">
                            <span class="d-dart-status-badge" style="background-color:${Helpers.escapeHtml(statusInfo.color)}">${Helpers.escapeHtml(statusInfo.display)}</span>
                            ${sowConfig ? `<span class="d-dart-sow-badge" data-toggle-target="${sowToggleId}" role="button" tabindex="0" aria-expanded="false" aria-controls="${sowToggleId}">SOW▼</span>` : '<span class="d-dart-sow-badge error">No SOW</span>'}
                        </div>
                    </div>
                    ${sowConfig ? HTMLGenerator.sowDetails(sowConfig, sowToggleId) : ''}
                    <div class="d-dart-id-row">
                        <a href="${smcUrl}" target="_blank" rel="noopener noreferrer" class="d-dart-id-item" title="Click to open in SMC"><div class="d-dart-id-content"><span class="d-dart-id-label">Order ID</span><span class="d-dart-id-value"><span class="d-dart-id-icon" aria-hidden="true">📋</span>${Helpers.escapeHtml(orderId)}</span></div></a>
                        <a href="${fmcVrUrl}" target="_blank" rel="noopener noreferrer" class="d-dart-id-item" title="Click to open in FMC"><div class="d-dart-id-content"><span class="d-dart-id-label">VR ID</span><span class="d-dart-id-value"><span class="d-dart-id-icon" aria-hidden="true">🚛</span>${Helpers.escapeHtml(vrId)}</span></div></a>
                        <a href="${fmcTourUrl}" target="_blank" rel="noopener noreferrer" class="d-dart-id-item" title="Click to open in FMC"><div class="d-dart-id-content"><span class="d-dart-id-label">Tour ID</span><span class="d-dart-id-value"><span class="d-dart-id-icon" aria-hidden="true">🔗</span>${Helpers.escapeHtml(tourId)}</span></div></a>
                    </div>
                    <div class="d-dart-lane-row">
                        <span class="d-dart-lane-origin">📍 ${Helpers.escapeHtml(originCode)}</span>
                        <div class="d-dart-lane-arrow-container"><div class="d-dart-lane-arrow-track"><span class="d-dart-arrow-moving" aria-hidden="true">→</span></div></div>
                        <span class="d-dart-lane-dest">${Helpers.escapeHtml(destCode)}</span>
                    </div>
                    <div class="d-dart-holds-row">
                        <span class="d-dart-holds-label">HOLDS:</span>
                        <span class="d-dart-hold-item ${holds.shipperExists ? 'has-hold' : 'no-hold'}">${holds.shipperExists ? '🟢' : '🔴'} Shipper: ${holds.shipperExists ? '$' + holds.shipperAmount.toFixed(2) : 'No'}</span>
                        <span class="d-dart-hold-item ${holds.receiverExists ? 'has-hold' : 'no-hold'}">${holds.receiverExists ? '🟢' : '🔴'} Receiver: ${holds.receiverExists ? '$' + holds.receiverAmount.toFixed(2) : 'No'}</span>
                    </div>
                </div>
            `;
        },

        getStopTimestamps(analysis, timezone) {
            const result = {
                arrival: { planned: '-', actual: '-', delay: null, delayClass: 'pending', borrowedVrId: null },
                departure: { planned: '-', actual: '-', delay: null, delayClass: 'pending', borrowedVrId: null }
            };
            if (!analysis?.fmcStopData?.timestamps) return result;
            const ts = analysis.fmcStopData.timestamps;
            if (ts.plannedYardArrival) result.arrival.planned = Helpers.formatTimeFromEpoch(ts.plannedYardArrival, timezone);
            if (ts.actualYardArrival) {
                result.arrival.actual = Helpers.formatTimeFromEpoch(ts.actualYardArrival, timezone);
                result.arrival.borrowedVrId = ts.actualYardArrivalSourceVrId;
                if (ts.plannedYardArrival) { const diff = DataHelpers.calculateTimeDiff(ts.plannedYardArrival, ts.actualYardArrival); result.arrival.delay = diff.minutes; result.arrival.delayClass = diff.minutes > 0 ? 'late' : 'early'; }
            }
            if (ts.plannedYardDeparture) result.departure.planned = Helpers.formatTimeFromEpoch(ts.plannedYardDeparture, timezone);
            if (ts.actualYardDeparture) {
                result.departure.actual = Helpers.formatTimeFromEpoch(ts.actualYardDeparture, timezone);
                result.departure.borrowedVrId = ts.actualYardDepartureSourceVrId;
                if (ts.plannedYardDeparture) { const diff = DataHelpers.calculateTimeDiff(ts.plannedYardDeparture, ts.actualYardDeparture); result.departure.delay = diff.minutes; result.departure.delayClass = diff.minutes > 0 ? 'late' : 'early'; }
            }
            return result;
        },

        timestampBox(label, data) {
            let delayText = '-'; let delayLabel = '-';
            if (data.delay !== null) {
                const absMinutes = Math.abs(data.delay);
                if (absMinutes >= 60) { const hours = Math.floor(absMinutes / 60); const mins = absMinutes % 60; delayText = `${hours}h ${mins}m`; }
                else delayText = `${absMinutes}m`;
                if (data.delay < 0) delayLabel = 'EARLY'; else if (data.delay > 0) delayLabel = 'DELAY'; else delayLabel = 'ON TIME';
            }
            const borrowedHtml = data.borrowedVrId ? Helpers.formatBorrowedIndicator(data.borrowedVrId) : '';
            return `
                <div class="d-dart-ts-box">
                    <div class="d-dart-ts-col planned"><div class="d-dart-ts-col-label">${Helpers.escapeHtml(label)} Planned</div><div class="d-dart-ts-col-value">${Helpers.escapeHtml(data.planned)}</div></div>
                    <div class="d-dart-ts-col actual"><div class="d-dart-ts-col-label">${Helpers.escapeHtml(label)} Actual</div><div class="d-dart-ts-col-value">${Helpers.escapeHtml(data.actual)}</div>${borrowedHtml}</div>
                    <div class="d-dart-ts-col delay ${data.delayClass}"><div class="d-dart-ts-col-label">${delayLabel}</div><div class="d-dart-ts-col-value">${delayText}</div></div>
                </div>
            `;
        },

        stopCard(stop, analysis, sowConfig, smcExecutionData) {
            const stopTypeInfo = analysis?.stopType || DataHelpers.formatStopType(stop?.stopActionType);
            const loadTypeInfo = DataHelpers.formatLoadType(stop?.loadingType);
            let timezone = 'America/Chicago';
            if (analysis?.fmcStopData) timezone = analysis.fmcStopData.timezone || timezone;
            else if (smcExecutionData) { timezone = stopTypeInfo.isPickup ? smcExecutionData.origin?.timezone : smcExecutionData.destination?.timezone; timezone = timezone || 'America/Chicago'; }
            const stopName = stop?.stopName || analysis?.stopName || `Stop ${(analysis?.stopIndex || 0) + 1}`;
            const fmcTimestamps = HTMLGenerator.getStopTimestamps(analysis, timezone);
            const borrowedWarning = analysis?.hasBorrowedTimestamps ? `<div class="d-dart-borrowed-warning">⚠️ Timestamp Issue — Data borrowed from another VR in this tour</div>` : '';
            const laneMismatchIndicator = analysis?.laneMismatch ? `<div class="d-dart-lane-mismatch-stop-indicator">🚫 Lane Mismatch — No charge for this stop</div>` : '';

            return `
                <div class="d-dart-stop-card ${analysis?.laneMismatch ? 'd-dart-stop-lane-mismatch' : ''}">
                    <div class="d-dart-stop-header">
                        <div class="d-dart-stop-title">Stop ${(analysis?.stopIndex || 0) + 1}: ${Helpers.escapeHtml(stopName)}</div>
                        <div class="d-dart-stop-badges">
                            <span class="d-dart-stop-badge ${stopTypeInfo.class}">${Helpers.escapeHtml(stopTypeInfo.label)}</span>
                            <span class="d-dart-stop-badge load-type">${loadTypeInfo.icon} ${Helpers.escapeHtml(loadTypeInfo.display)}</span>
                            ${analysis?.laneMismatch ? '<span class="d-dart-stop-badge lane-mismatch">🚫 LANE MISMATCH</span>' : ''}
                        </div>
                    </div>
                    ${laneMismatchIndicator}
                    <div class="d-dart-timestamp-boxes">
                        ${HTMLGenerator.timestampBox('Arrival', fmcTimestamps.arrival)}
                        ${HTMLGenerator.timestampBox('Departure', fmcTimestamps.departure)}
                    </div>
                    ${borrowedWarning}
                </div>
            `;
        },

        batchReportTable(data) {
            if (!data || data.length === 0) return '<div class="d-dart-empty">No results to display</div>';
            const isReview = ModeManager.isReviewMode();

            const stats = {
                recovered: data.filter(d => d.action === ActionDisplayConfig.RECOVERED.reportTerm || d.action === 'Would Recover').length,
                chargesAdded: data.filter(d => d.action === ActionDisplayConfig.CHARGE_ADDED.reportTerm || d.action === 'Would Add Charge').length,
                holdsReleased: data.filter(d => d.action === ActionDisplayConfig.HOLD_RELEASED.reportTerm || d.action === 'Would Release Hold').length,
                holdsCreated: data.filter(d => d.action === ActionDisplayConfig.HOLD_CREATED.reportTerm || d.action === 'Would Create Hold').length,
                analysisOnly: data.filter(d => d.action === ActionDisplayConfig.ANALYSIS_ONLY.reportTerm).length,
                pending: data.filter(d => d.action === ActionDisplayConfig.PENDING.reportTerm || d.action === 'Would Charge (Approval Required)').length,
                errors: data.filter(d => d.status === 'Error').length,
                mismatches: data.filter(d => d.action && d.action.includes('Charge Mismatch')).length,
                correct: data.filter(d => d.action && d.action.includes('Charge Correct')).length,
                laneMismatch: data.filter(d => d.action === ActionDisplayConfig.LANE_MISMATCH.reportTerm || d.action === 'Would Release (Lane Mismatch)').length
            };

            const rowsHtml = data.map(row => {
                let statusClass = '';
                if (row.status === 'Completed') statusClass = 'success';
                else if (row.status === 'Review') statusClass = 'analysis';
                else if (row.status === 'Pending') statusClass = 'pending';
                else if (row.status === 'Error') statusClass = 'error';
                else if (row.status === 'Info') statusClass = 'analysis';

                let actionClass = '';
                if (row.action === ActionDisplayConfig.RECOVERED.reportTerm || row.action === 'Would Recover') actionClass = 'recovered';
                if (row.action === ActionDisplayConfig.ANALYSIS_ONLY.reportTerm) actionClass = 'analysis';
                if (row.action === ActionDisplayConfig.HOLD_CREATED.reportTerm || row.action === 'Would Create Hold') actionClass = 'hold-created';
                if (row.action && row.action.includes('Charge Mismatch')) actionClass = 'mismatch';
                if (row.action && row.action.includes('Charge Correct')) actionClass = 'correct';
                if (row.action === ActionDisplayConfig.LANE_MISMATCH.reportTerm || row.action === 'Would Release (Lane Mismatch)') actionClass = 'lane-mismatch-row';

                return `<tr class="${statusClass} ${actionClass}"><td class="d-dart-order-cell">${Helpers.escapeHtml(row.orderId)}</td><td>${Helpers.escapeHtml(Helpers.truncateText(row.shipper, 15))}</td><td>${Helpers.escapeHtml(row.action)}</td><td>${Helpers.escapeHtml(row.amount)}</td><td><span class="d-dart-status-badge ${statusClass}">${Helpers.escapeHtml(row.status)}</span></td></tr>`;
            }).join('');

            const reviewBanner = isReview ? `<div class="d-dart-review-banner">${Messages.INFO.REVIEW_MODE_BANNER}</div>` : '';

            return `
                ${reviewBanner}
                <div class="d-dart-batch-summary">
                    <div class="d-dart-batch-stat recovered">🎯 ${stats.recovered} Rec</div>
                    <div class="d-dart-batch-stat success">✅ ${stats.chargesAdded} Add</div>
                    <div class="d-dart-batch-stat released">✅ ${stats.holdsReleased} Rel</div>
                    <div class="d-dart-batch-stat hold-created">📋 ${stats.holdsCreated} Hold</div>
                    <div class="d-dart-batch-stat analysis">📊 ${stats.analysisOnly} Info</div>
                    <div class="d-dart-batch-stat pending">⏳ ${stats.pending} Pend</div>
                    <div class="d-dart-batch-stat lane-mismatch-stat">🚫 ${stats.laneMismatch} Lane</div>
                    <div class="d-dart-batch-stat mismatch-stat">⚠️ ${stats.mismatches} Mis</div>
                    <div class="d-dart-batch-stat error">❌ ${stats.errors} Err</div>
                </div>
                <div class="d-dart-batch-table-container">
                    <table class="d-dart-batch-table"><thead><tr><th>Order ID</th><th>Shipper</th><th>Action</th><th>Amount</th><th>Status</th></tr></thead><tbody>${rowsHtml}</tbody></table>
                </div>
                <div class="d-dart-download-buttons">
                    <button class="d-dart-download-btn" id="d-dart-download-csv">📥 CSV</button>
                    <button class="d-dart-download-btn" id="d-dart-download-txt">📄 TXT</button>
                </div>
            `;
        },

        enhancedBatchProgress(totalOrders, totalChunks) {
            return `
                <div class="d-dart-enhanced-progress">
                    <div class="d-dart-progress-header"><div class="d-dart-progress-title">🔄 Enterprise Batch Processing</div>
                        <div class="d-dart-progress-controls">
                            <button class="d-dart-control-btn pause" id="d-dart-pause-btn" title="Pause" aria-label="Pause processing">⏸️</button>
                            <button class="d-dart-control-btn resume" id="d-dart-resume-btn" title="Resume" style="display:none" aria-label="Resume processing">▶️</button>
                            <button class="d-dart-control-btn cancel" id="d-dart-cancel-btn" title="Cancel" aria-label="Cancel processing">⏹️</button>
                        </div>
                    </div>
                    <div class="d-dart-progress-stats">
                        <div class="d-dart-stat"><span class="d-dart-stat-value" id="d-dart-stat-processed">0</span><span class="d-dart-stat-label">Processed</span></div>
                        <div class="d-dart-stat success"><span class="d-dart-stat-value" id="d-dart-stat-success">0</span><span class="d-dart-stat-label">Success</span></div>
                        <div class="d-dart-stat error"><span class="d-dart-stat-value" id="d-dart-stat-failed">0</span><span class="d-dart-stat-label">Failed</span></div>
                        <div class="d-dart-stat"><span class="d-dart-stat-value" id="d-dart-stat-remaining">${totalOrders}</span><span class="d-dart-stat-label">Remaining</span></div>
                    </div>
                    <div class="d-dart-progress-bar-container"><div class="d-dart-progress-bar" id="d-dart-progress-bar" style="width: 0%" role="progressbar" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100"></div></div>
                    <div class="d-dart-progress-info"><div class="d-dart-progress-status" id="d-dart-progress-status" aria-live="polite">Initializing...</div><div class="d-dart-progress-eta" id="d-dart-progress-eta"></div></div>
                    <div class="d-dart-progress-footer"><div class="d-dart-progress-chunk">Chunk: <span id="d-dart-chunk-info">0/${totalChunks}</span></div><div class="d-dart-progress-token">Token: <span id="d-dart-token-status" class="token-ok">Ready</span></div></div>
                </div>
            `;
        },

        sowErrorDisplay(errorMessage, isAuthError = false) {
            return `
                <div class="d-dart-sow-error" role="alert">
                    <div class="d-dart-sow-error-icon" aria-hidden="true">${isAuthError ? '🔐' : '❌'}</div>
                    <div class="d-dart-sow-error-title">${isAuthError ? 'SharePoint Login Required' : 'SOW Server Unavailable'}</div>
                    <div class="d-dart-sow-error-message">${Helpers.escapeHtml(errorMessage)}</div>
                    ${isAuthError ? `<div class="d-dart-sow-error-instructions"><p>1. Click the button below to open SharePoint</p><p>2. Login with your credentials</p><p>3. Come back and click "Retry"</p></div><a href="${CONFIG.SHAREPOINT.SITE_URL}" target="_blank" rel="noopener noreferrer" class="d-dart-sow-login-btn">🔗 Open SharePoint</a>` : ''}
                    <button class="d-dart-sow-retry-btn" id="d-dart-sow-retry">🔄 Retry</button>
                </div>
            `;
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 26: REPORT GENERATOR (lane mismatch in reports)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const ReportGenerator = {
        generateCSV(data) {
            const headers = ['Order ID', 'Shipper', 'Action', 'Amount', 'Status', 'Notes'];
            const rows = data.map(row => [row.orderId, `"${String(row.shipper || '').replace(/"/g, '""')}"`, row.action, row.amount, row.status, `"${String(row.notes || '').replace(/"/g, '""')}"`].join(','));
            return [headers.join(','), ...rows].join('\n');
        },

        generateTXT(data) {
            const now = new Date().toLocaleString();
            const separator = '═'.repeat(79); const thinSeparator = '─'.repeat(79);
            const isReview = ModeManager.isReviewMode();
            const stats = {
                recovered: data.filter(d => d.action === ActionDisplayConfig.RECOVERED.reportTerm || d.action === 'Would Recover').length,
                chargesAdded: data.filter(d => d.action === ActionDisplayConfig.CHARGE_ADDED.reportTerm || d.action === 'Would Add Charge').length,
                holdsReleased: data.filter(d => d.action === ActionDisplayConfig.HOLD_RELEASED.reportTerm || d.action === 'Would Release Hold').length,
                holdsCreated: data.filter(d => d.action === ActionDisplayConfig.HOLD_CREATED.reportTerm || d.action === 'Would Create Hold').length,
                analysisOnly: data.filter(d => d.action === ActionDisplayConfig.ANALYSIS_ONLY.reportTerm).length,
                pending: data.filter(d => d.action === ActionDisplayConfig.PENDING.reportTerm || d.action === 'Would Charge (Approval Required)').length,
                noAction: data.filter(d => d.action === ActionDisplayConfig.NO_ACTION.reportTerm).length,
                errors: data.filter(d => d.status === 'Error').length,
                mismatches: data.filter(d => d.action && d.action.includes('Charge Mismatch')).length,
                correct: data.filter(d => d.action && d.action.includes('Charge Correct')).length,
                laneMismatch: data.filter(d => d.action === ActionDisplayConfig.LANE_MISMATCH.reportTerm || d.action === 'Would Release (Lane Mismatch)').length
            };

            let report = `
${separator}
              D-DART Enterprise Batch Report v${CONFIG.VERSION}
${separator}
 Generated: ${now}
 Mode: ${isReview ? 'REVIEW (No actions taken)' : 'ACTION'}
 Total Orders Processed: ${data.length}
${separator}

SUMMARY:
${thinSeparator}
  🎯 Recovered:           ${stats.recovered}
  ✅ Charges Added:       ${stats.chargesAdded}
  ✅ Holds Released:      ${stats.holdsReleased}
  📋 Holds Created:       ${stats.holdsCreated}
  📊 Analysis Only:       ${stats.analysisOnly}
  ⏳ Pending:             ${stats.pending}
  ⏸️ No Action:           ${stats.noAction}
  🚫 Lane Mismatch:       ${stats.laneMismatch}
  ⚠️ Charge Mismatches:   ${stats.mismatches}
  ✅ Charges Correct:     ${stats.correct}
  ❌ Errors:              ${stats.errors}

${separator}
DETAILED RESULTS:
${thinSeparator}
`;
            data.forEach((row, index) => {
                report += `\n${index + 1}. Order: ${row.orderId}\n   Shipper: ${row.shipper}\n   Action: ${row.action}\n   Amount: ${row.amount}\n   Status: ${row.status}\n   ${row.notes ? `Notes: ${row.notes}` : ''}\n${thinSeparator}`;
            });
            report += `\n\n${separator}\n                              End of Report\n${separator}\n`;
            return report.trim();
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 27: APPROVAL POPUP
     * ═══════════════════════════════════════════════════════════════════════════ */

    const ApprovalPopup = (() => {
        let currentPopup = null; let timeoutId = null; let countdownInterval = null;
        let resolveCallback = null; let popupElements = null; let mutationObserver = null;

        const cleanup = () => {
            if (countdownInterval) { clearInterval(countdownInterval); countdownInterval = null; }
            if (timeoutId) { clearTimeout(timeoutId); timeoutId = null; }
            if (mutationObserver) { mutationObserver.disconnect(); mutationObserver = null; }
            if (currentPopup) { currentPopup.remove(); currentPopup = null; }
            popupElements = null;
        };

        const stopCountdown = () => { if (countdownInterval) { clearInterval(countdownInterval); countdownInterval = null; } if (timeoutId) { clearTimeout(timeoutId); timeoutId = null; } };

        const handleYes = (authNumber) => { stopCountdown(); cleanup(); Logger.info('Approval: User approved', { authNumber: authNumber ? 'provided' : 'not required' }); if (resolveCallback) { resolveCallback({ decision: 'YES', authorizationNumber: authNumber }); resolveCallback = null; } };
        const handleNo = () => { stopCountdown(); cleanup(); Logger.info('Approval: User declined'); if (resolveCallback) { resolveCallback({ decision: 'NO' }); resolveCallback = null; } };
        const handleSkip = () => { stopCountdown(); cleanup(); Logger.info('Approval: User skipped'); if (resolveCallback) { resolveCallback({ decision: 'SKIP' }); resolveCallback = null; } };
        const handleTimeout = () => { stopCountdown(); cleanup(); Logger.info('Approval: Timeout'); if (resolveCallback) { resolveCallback({ decision: 'TIMEOUT' }); resolveCallback = null; } };

        const startCountdown = () => {
            let secondsLeft = Math.floor(CONFIG.APPROVAL.TIMEOUT / 1000);
            countdownInterval = setInterval(() => {
                secondsLeft--;
                if (popupElements?.timer) {
                    popupElements.timer.textContent = `⏱️ ${secondsLeft}s`;
                    popupElements.timer.classList.remove('warning', 'critical');
                    if (secondsLeft <= CONFIG.APPROVAL.WARNING_TIME) popupElements.timer.classList.add('warning');
                    if (secondsLeft <= CONFIG.APPROVAL.CRITICAL_TIME) popupElements.timer.classList.add('critical');
                }
                if (secondsLeft <= 0) handleTimeout();
            }, CONFIG.APPROVAL.COUNTDOWN_INTERVAL);
            timeoutId = setTimeout(() => { handleTimeout(); }, CONFIG.APPROVAL.TIMEOUT);
        };

        const showAuthorizationInput = (orderData, totalCharge) => {
            stopCountdown();
            const popup = currentPopup?.querySelector('.d-dart-approval-popup');
            if (!popup) return;
            popup.innerHTML = `
                <div class="d-dart-approval-header"><span class="d-dart-approval-title">✅ ENTER AUTHORIZATION</span></div>
                <div class="d-dart-approval-body">
                    <div class="d-dart-approval-order-id"><span class="d-dart-approval-label">Order ID:</span><span class="d-dart-approval-value">${Helpers.escapeHtml(orderData?.orderId || 'Unknown')}</span></div>
                    <div class="d-dart-approval-charge-confirm"><span class="d-dart-approval-label">Charge Amount:</span><span class="d-dart-approval-value">${Helpers.formatCurrency(totalCharge)}</span></div>
                    <div class="d-dart-approval-auth-section">
                        <label class="d-dart-approval-auth-label" for="d-dart-approval-auth-input">Please enter the Authorization Number:</label>
                        <input type="text" class="d-dart-approval-auth-input" id="d-dart-approval-auth-input" placeholder="Enter authorization number" autocomplete="off"><div class="d-dart-approval-auth-error" id="d-dart-approval-auth-error" role="alert"></div>
                    </div>
                </div>
                <div class="d-dart-approval-buttons"><button class="d-dart-approval-btn submit" id="d-dart-approval-submit">SUBMIT</button><button class="d-dart-approval-btn cancel" id="d-dart-approval-cancel">CANCEL</button></div>
            `;
            popupElements = { authInput: document.getElementById('d-dart-approval-auth-input'), authError: document.getElementById('d-dart-approval-auth-error'), submitBtn: document.getElementById('d-dart-approval-submit'), cancelBtn: document.getElementById('d-dart-approval-cancel') };
            setTimeout(() => { popupElements?.authInput?.focus(); }, 100);
            const handleAuthSubmit = () => { const rawAuthNumber = popupElements?.authInput?.value; const authNumber = Validator.sanitizeAuthNumber(rawAuthNumber); if (!authNumber) { if (popupElements?.authError) popupElements.authError.textContent = Messages.ERRORS.AUTH_NUMBER_REQUIRED; popupElements?.authInput?.classList.add('error'); popupElements?.authInput?.focus(); return; } handleYes(authNumber); };
            popupElements?.submitBtn?.addEventListener('click', handleAuthSubmit);
            popupElements?.cancelBtn?.addEventListener('click', handleSkip);
            popupElements?.authInput?.addEventListener('keypress', (e) => { if (e.key === 'Enter') handleAuthSubmit(); });
            popupElements?.authInput?.addEventListener('input', () => { if (popupElements?.authError) popupElements.authError.textContent = ''; popupElements?.authInput?.classList.remove('error'); });
        };

        const createPopup = (orderData, totalCharge, stopDetails) => {
            cleanup();
            const sowConfig = orderData?.sowConfig;
            const requiresAuth = sowConfig?.authNumberRequired || false;
            const overlay = document.createElement('div');
            overlay.id = 'd-dart-approval-overlay';
            overlay.setAttribute('role', 'dialog'); overlay.setAttribute('aria-modal', 'true'); overlay.setAttribute('aria-labelledby', 'd-dart-approval-title');
            overlay.innerHTML = `
                <div class="d-dart-approval-popup">
                    <div class="d-dart-approval-header"><span class="d-dart-approval-title" id="d-dart-approval-title">⚠️ APPROVAL REQUIRED</span><span class="d-dart-approval-timer" id="d-dart-approval-timer">⏱️ ${Math.floor(CONFIG.APPROVAL.TIMEOUT / 1000)}s</span></div>
                    <div class="d-dart-approval-body">
                        <div class="d-dart-approval-order-id"><span class="d-dart-approval-label">Order ID:</span><span class="d-dart-approval-value">${Helpers.escapeHtml(orderData?.orderId || 'Unknown')}</span></div>
                        <div class="d-dart-approval-shipper"><span class="d-dart-approval-label">Shipper:</span><span class="d-dart-approval-value">${Helpers.escapeHtml(orderData?.shipperName || 'Unknown')}</span></div>
                        <div class="d-dart-approval-charge-info">
                            <div class="d-dart-approval-charge-title">Possible detention charge detected:</div>
                            <div class="d-dart-approval-charge-details">${stopDetails.map(s => `<div class="d-dart-approval-stop-line">📍 ${Helpers.escapeHtml(s.stopType)}: <strong>${Helpers.formatCurrency(s.charge)}</strong></div>`).join('')}</div>
                            <div class="d-dart-approval-total"><span class="d-dart-approval-total-label">Total Amount:</span><span class="d-dart-approval-total-value">${Helpers.formatCurrency(totalCharge)}</span></div>
                        </div>
                        <div class="d-dart-approval-question">Do you approve this charge?</div>
                    </div>
                    <div class="d-dart-approval-buttons">
                        <button class="d-dart-approval-btn yes" id="d-dart-approval-yes" aria-label="${Messages.ACCESSIBILITY.APPROVE_CHARGE}">✅ YES</button>
                        <button class="d-dart-approval-btn no" id="d-dart-approval-no" aria-label="${Messages.ACCESSIBILITY.DECLINE_CHARGE}">❌ NO</button>
                        <button class="d-dart-approval-btn skip" id="d-dart-approval-skip" aria-label="${Messages.ACCESSIBILITY.SKIP_ORDER}">⏭️ SKIP</button>
                    </div>
                </div>
            `;
            document.body.appendChild(overlay);
            currentPopup = overlay;
            popupElements = { timer: document.getElementById('d-dart-approval-timer'), yesBtn: document.getElementById('d-dart-approval-yes'), noBtn: document.getElementById('d-dart-approval-no'), skipBtn: document.getElementById('d-dart-approval-skip') };

            mutationObserver = new MutationObserver((mutations) => {
                for (const mutation of mutations) {
                    for (const removedNode of mutation.removedNodes) {
                        if (removedNode === currentPopup || removedNode.contains?.(currentPopup)) {
                            Logger.warn('Approval popup removed externally, cleaning up');
                            stopCountdown();
                            if (resolveCallback) { resolveCallback({ decision: 'SKIP' }); resolveCallback = null; }
                            mutationObserver?.disconnect(); mutationObserver = null; currentPopup = null; popupElements = null;
                            return;
                        }
                    }
                }
            });
            mutationObserver.observe(document.body, { childList: true });

            popupElements.yesBtn?.addEventListener('click', () => { if (requiresAuth) showAuthorizationInput(orderData, totalCharge); else handleYes(null); });
            popupElements.noBtn?.addEventListener('click', handleNo);
            popupElements.skipBtn?.addEventListener('click', handleSkip);
            overlay.addEventListener('keydown', (e) => { if (e.key === 'Escape') handleSkip(); });
            popupElements.yesBtn?.focus();
        };

        return {
            show(orderData, totalCharge, stopDetails) { return new Promise((resolve) => { resolveCallback = resolve; createPopup(orderData, totalCharge, stopDetails); startCountdown(); }); },
            cleanup
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 28: BATCH PROCESSOR (with lane validation + all fixes)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const BatchProcessor = (() => {
        let batchState = BatchState.IDLE;
        let startTime = null;
        let lastUIUpdate = 0;
        // Fix #11: cached progress DOM refs
        let progressDOM = null;

        const initializeOrderData = (orderId) => ({
            orderId, viewData: null, fullData: null, smcExecutionData: null, fmcData: null,
            fmcTimestamps: null, analysisResults: [], shipperName: 'Unknown', sowConfig: null,
            laneValidation: null, matchingVR: null
        });

        const analyzeStops = (orderData) => {
            const stops = orderData?.viewData?.stops || [];
            const pricing = orderData?.viewData?.shipperPricing?.pricing || [];
            const orderStatus = orderData?.viewData?.orderExecutionStatus || 'UNKNOWN';
            const holdInfo = DataHelpers.findDetentionHolds(pricing);
            // Pass lane validation to every stop (whole-order decision)
            const laneMismatchInfo = (orderData.laneValidation && !orderData.laneValidation.isMatch) ? orderData.laneValidation : null;

            for (let index = 0; index < stops.length; index++) {
                const stop = stops[index];
                const stopTypeInfo = DataHelpers.formatStopType(stop?.stopActionType);
                let fmcStopTimestamps = null;
                if (orderData.fmcTimestamps) fmcStopTimestamps = stopTypeInfo.isPickup ? orderData.fmcTimestamps.origin : orderData.fmcTimestamps.destination;

                const analysis = DetentionAnalyzer.analyze(stop, fmcStopTimestamps, orderData.sowConfig, orderStatus, holdInfo, index, orderData.shipperName, laneMismatchInfo);
                analysis.stopIndex = index; analysis.stopName = stop?.stopName || `Stop ${index + 1}`;
                analysis.stopType = stopTypeInfo; analysis.stop = stop; analysis.fmcStopData = fmcStopTimestamps;

                if (fmcStopTimestamps?.timestamps) {
                    if (fmcStopTimestamps.timestamps.actualYardArrivalSourceVrId || fmcStopTimestamps.timestamps.actualYardDepartureSourceVrId) analysis.hasBorrowedTimestamps = true;
                }

                orderData.analysisResults.push(analysis);
            }
        };

        const addToBatchReport = (orderData) => {
            const isReview = ModeManager.isReviewMode();
            const entry = { orderId: orderData?.orderId || 'Unknown', shipper: orderData?.shipperName || 'Unknown', action: ActionDisplayConfig.NO_ACTION.reportTerm, amount: '$0.00', status: isReview ? 'Review' : 'Completed', notes: '' };

            // Check for lane mismatch first
            const hasLaneMismatch = (orderData?.analysisResults || []).some(a => a?.laneMismatch);
            if (hasLaneMismatch) {
                if (isReview) {
                    const hasHolds = (orderData?.analysisResults || []).some(a => a?.hasHold);
                    entry.action = hasHolds ? 'Would Release (Lane Mismatch)' : ActionDisplayConfig.LANE_MISMATCH.reportTerm;
                    entry.amount = '$0.00';
                    entry.status = 'Review';
                    const lv = orderData?.laneValidation;
                    if (lv) entry.notes = `SMC: ${lv.smcOrigin}→${lv.smcDest} | FMC: ${lv.fmcOrigin}→${lv.fmcDest}`;
                } else {
                    entry.action = ActionDisplayConfig.LANE_MISMATCH.reportTerm;
                    entry.amount = '$0.00';
                    entry.status = 'Completed';
                    const lv = orderData?.laneValidation;
                    if (lv) entry.notes = `SMC: ${lv.smcOrigin}→${lv.smcDest} | FMC: ${lv.fmcOrigin}→${lv.fmcDest}`;
                }
                AppState.addBatchReportEntry(entry);
                return;
            }

            let totalCharge = 0;
            let hasError = false; let hasPending = false;
            let hasChargeAdded = false; let hasRecovered = false;
            let hasHoldReleased = false; let hasHoldCreated = false;
            let hasAnalysisOnly = false; let hasApprovalRequired = false;
            let hasMismatch = false; let mismatchDetails = null;
            let hasCorrectCharge = false;
            const notesList = [];

            for (const analysis of (orderData?.analysisResults || [])) {
                if (analysis?.processedAmount > 0) totalCharge += analysis.processedAmount;
                else if (analysis?.charge > 0 && analysis?.processed) totalCharge += analysis.charge;
                else if (analysis?.charge > 0 && (analysis?.action === ActionType.ANALYSIS_ONLY || analysis?.action === ActionType.CREATE_HOLD_ONLY)) totalCharge += analysis.charge;
                else if (analysis?.charge > 0 && (analysis?.action === ActionType.ADD_CHARGE_UPDATE || analysis?.action === ActionType.ADD_CHARGE_CREATE)) totalCharge += analysis.charge;
                else if (analysis?.charge > 0 && analysis?.action === ActionType.PENDING_APPROVAL) totalCharge += analysis.charge;

                if (analysis?.processError) { hasError = true; notesList.push(analysis.processError); }
                else if (analysis?.action === ActionType.PENDING || analysis?.action === ActionType.ERROR) { hasPending = true; notesList.push(analysis.actionText || 'Pending'); }

                if (analysis?.processed) {
                    switch (analysis.processedAction) {
                        case 'updated': hasChargeAdded = true; break;
                        case 'created': hasRecovered = true; break;
                        case 'released': case 'released_lane_mismatch': hasHoldReleased = true; break;
                        case 'hold_created': hasHoldCreated = true; break;
                        case 'skipped': case 'timeout': hasPending = true; break;
                        case 'analysis_only': hasAnalysisOnly = true; break;
                    }
                }

                if (!analysis?.processed && !analysis?.processError) {
                    if (analysis?.action === ActionType.CREATE_HOLD_ONLY) hasHoldCreated = true;
                    if (analysis?.action === ActionType.ANALYSIS_ONLY) hasAnalysisOnly = true;
                    if (analysis?.action === ActionType.PENDING_APPROVAL) hasApprovalRequired = true;
                    if (analysis?.action === ActionType.ADD_CHARGE_UPDATE) hasChargeAdded = true;
                    if (analysis?.action === ActionType.ADD_CHARGE_CREATE) hasRecovered = true;
                    if (analysis?.action === ActionType.RELEASE_HOLD) hasHoldReleased = true;
                }

                if (analysis?.mismatchType) {
                    if (analysis.mismatchType === 'CORRECT') { hasCorrectCharge = true; }
                    else { hasMismatch = true; mismatchDetails = analysis; }
                }

                if (analysis?.type === ResultType.CHARGE_EXISTS) notesList.push(`Existing: $${(analysis.charge || 0).toFixed(2)}`);
                else if (analysis?.type === ResultType.DRIVER_LATE) notesList.push('Driver late');
                else if (analysis?.type === ResultType.NO_DETENTION_DROP_HOOK) notesList.push('Not eligible');
                else if (analysis?.type === ResultType.BELOW_MINIMUM_THRESHOLD) notesList.push('Below minimum');

                if (analysis?.hasBorrowedTimestamps) notesList.push(Messages.INFO.BORROWED_TIMESTAMP_NOTE);
                if (analysis?.approvalNote) notesList.push(analysis.approvalNote);
            }

            if (hasMismatch && mismatchDetails) {
                const statusTag = mismatchDetails.orderStatusGroup === 'invoiced' ? ' [Invoiced]' : mismatchDetails.orderStatusGroup === 'paid' ? ' [Paid]' : '';
                switch (mismatchDetails.mismatchType) {
                    case 'UNDER':
                        entry.action = `Charge Mismatch (Under)${statusTag}`;
                        entry.amount = `SMC: $${mismatchDetails.existingAmount.toFixed(2)} / Correct: $${mismatchDetails.calculatedAmount.toFixed(2)}`;
                        break;
                    case 'OVER':
                        entry.action = `Charge Mismatch (Over)${statusTag}`;
                        entry.amount = `SMC: $${mismatchDetails.existingAmount.toFixed(2)} / Correct: $${mismatchDetails.calculatedAmount.toFixed(2)}`;
                        break;
                    case 'SHOULD_NOT_EXIST':
                        entry.action = `Charge Mismatch (Should Not Exist)${statusTag}`;
                        entry.amount = `SMC: $${mismatchDetails.existingAmount.toFixed(2)} / Correct: $0.00`;
                        notesList.push(mismatchDetails.noChargeReason || 'N/A');
                        break;
                }
                entry.status = 'Review';
            } else if (hasCorrectCharge && !hasMismatch) {
                const correctAnalysis = (orderData?.analysisResults || []).find(a => a?.mismatchType === 'CORRECT');
                const statusTag = correctAnalysis?.orderStatusGroup === 'invoiced' ? ' [Invoiced]' : correctAnalysis?.orderStatusGroup === 'paid' ? ' [Paid]' : '';
                entry.action = `Charge Correct${statusTag}`;
                entry.amount = correctAnalysis ? `$${correctAnalysis.existingAmount.toFixed(2)}` : '-';
                entry.status = 'Review';
            } else if (isReview) {
                if (hasRecovered) { entry.action = 'Would Recover'; entry.amount = Helpers.formatCurrency(totalCharge); }
                else if (hasChargeAdded) { entry.action = 'Would Add Charge'; entry.amount = Helpers.formatCurrency(totalCharge); }
                else if (hasHoldReleased) { entry.action = 'Would Release Hold'; entry.amount = '$0.00'; }
                else if (hasHoldCreated) { entry.action = 'Would Create Hold'; entry.amount = totalCharge > 0 ? Helpers.formatCurrency(totalCharge) + ' (chargeable)' : '-'; }
                else if (hasApprovalRequired) { entry.action = 'Would Charge (Approval Required)'; entry.amount = totalCharge > 0 ? Helpers.formatCurrency(totalCharge) : '-'; }
                else if (hasAnalysisOnly) { entry.action = ActionDisplayConfig.ANALYSIS_ONLY.reportTerm; entry.amount = totalCharge > 0 ? Helpers.formatCurrency(totalCharge) : '-'; }
                else if (hasPending) { entry.action = ActionDisplayConfig.PENDING.reportTerm; entry.amount = totalCharge > 0 ? Helpers.formatCurrency(totalCharge) : '-'; }
                entry.status = 'Review';
            } else {
                if (hasRecovered) { entry.action = ActionDisplayConfig.RECOVERED.reportTerm; entry.amount = Helpers.formatCurrency(totalCharge); }
                else if (hasChargeAdded) { entry.action = ActionDisplayConfig.CHARGE_ADDED.reportTerm; entry.amount = Helpers.formatCurrency(totalCharge); }
                else if (hasHoldReleased) { entry.action = ActionDisplayConfig.HOLD_RELEASED.reportTerm; entry.amount = '$0.00'; }
                else if (hasHoldCreated) { entry.action = ActionDisplayConfig.HOLD_CREATED.reportTerm; entry.amount = totalCharge > 0 ? Helpers.formatCurrency(totalCharge) + ' (chargeable)' : '-'; entry.status = 'Info'; notesList.push('$0 hold created for manual review'); }
                else if (hasAnalysisOnly) { entry.action = ActionDisplayConfig.ANALYSIS_ONLY.reportTerm; entry.amount = totalCharge > 0 ? Helpers.formatCurrency(totalCharge) : '-'; entry.status = 'Info'; notesList.push('Hold already exists'); }
                else if (hasPending) { entry.action = ActionDisplayConfig.PENDING.reportTerm; entry.amount = totalCharge > 0 ? Helpers.formatCurrency(totalCharge) : '-'; }
            }

            if (hasError) entry.status = 'Error';
            else if (hasPending && !isReview) entry.status = 'Pending';

            entry.notes = notesList.join('; ');
            AppState.addBatchReportEntry(entry);
        };

        const ensureValidToken = async () => {
            if (ModeManager.isReviewMode()) {
                const remaining = TokenManager.getRemainingSeconds();
                if (remaining <= 0) {
                    Logger.info('Review Mode: Session may be expired, attempting token refresh as health check');
                    const success = await TokenManager.ensure();
                    if (!success) {
                        Logger.warn('Review Mode: Session health check failed — reads may fail');
                        UIController.showToast(Messages.INFO.SESSION_WARNING, 'warning');
                    }
                }
                return;
            }
            const remaining = TokenManager.getRemainingSeconds();
            if (remaining < CONFIG.TOKEN.WARNING_THRESHOLD) {
                UIController.updateBatchStatus(Messages.INFO.TOKEN_REFRESHING);
                const success = await TokenManager.ensure();
                if (!success) throw ErrorHandler.create(ErrorType.AUTH, Messages.ERRORS.TOKEN_MISSING, { action: 'ensureToken' });
            }
        };

        const executeActionsWithRetry = async (orderData, maxRetries = 2) => {
            for (let attempt = 0; attempt <= maxRetries; attempt++) {
                try { await executeActions(orderData); return; }
                catch (error) {
                    if (error.type === ErrorType.BUSINESS && error.message?.includes('Version conflict')) {
                        if (attempt < maxRetries) { Logger.warn(`Version conflict, refetching order (attempt ${attempt + 1})`); orderData.fullData = await SMCApiService.fetchOrderFull(orderData.orderId); continue; }
                    }
                    throw error;
                }
            }
        };

        const executeActions = async (orderData) => {
            const orderId = orderData?.orderId;
            if (!orderId) return;

            let chargeAdded = false;
            let releaseProcessed = false;
            let hasReleasedWithBorrowed = false;
            let hasReleasedWithLaneMismatch = false;

            const releaseActions = (orderData.analysisResults || []).filter(a => a?.action === ActionType.RELEASE_HOLD);
            const updateActions = (orderData.analysisResults || []).filter(a => a?.action === ActionType.ADD_CHARGE_UPDATE);
            const createActions = (orderData.analysisResults || []).filter(a => a?.action === ActionType.ADD_CHARGE_CREATE);
            const createHoldOnlyActions = (orderData.analysisResults || []).filter(a => a?.action === ActionType.CREATE_HOLD_ONLY);

            const needsFullData = releaseActions.length > 0 || updateActions.length > 0 || createActions.length > 0 || createHoldOnlyActions.length > 0;
            if (!needsFullData) return;

            if (!orderData.fullData) orderData.fullData = await SMCApiService.fetchOrderFull(orderId);
            let currentVersion = orderData.fullData?.orderId?.version;

            if (releaseActions.length > 0 || updateActions.length > 0) {
                const pricing = [...(orderData.fullData?.orderDetails?.shipperPricing?.pricing || [])];
                let modified = false;

                for (const analysis of updateActions) {
                    if (analysis?.holdCode) {
                        const itemIndex = pricing.findIndex(p => p?.pricingCode === analysis.holdCode);
                        if (itemIndex !== -1) {
                            pricing[itemIndex] = { ...pricing[itemIndex], price: { ...pricing[itemIndex].price, value: analysis.charge } };
                            analysis.processed = true; analysis.processedAction = 'updated'; analysis.processedAmount = analysis.charge;
                            modified = true; chargeAdded = true;
                        } else {
                            Logger.warn(`Hold code ${analysis.holdCode} not found in pricing for ${orderId}`);
                            analysis.processError = Messages.ERRORS.HOLD_CODE_NOT_FOUND(analysis.holdCode, orderId);
                        }
                    } else {
                        Logger.warn(`Update action has no holdCode for ${orderId}`);
                        analysis.processError = Messages.ERRORS.HOLD_CODE_MISSING;
                    }
                }

                for (const analysis of releaseActions) {
                    if (analysis?.holdCode) {
                        const itemIndex = pricing.findIndex(p => p?.pricingCode === analysis.holdCode);
                        if (itemIndex !== -1) {
                            pricing.splice(itemIndex, 1);
                            analysis.processed = true;
                            analysis.processedAction = analysis.laneMismatch ? 'released_lane_mismatch' : 'released';
                            modified = true; releaseProcessed = true;
                            if (analysis.hasBorrowedTimestamps) hasReleasedWithBorrowed = true;
                            if (analysis.laneMismatch) hasReleasedWithLaneMismatch = true;
                        } else {
                            Logger.warn(`Hold code ${analysis.holdCode} not found for release in ${orderId}`);
                            analysis.processError = Messages.ERRORS.HOLD_CODE_NOT_FOUND(analysis.holdCode, orderId);
                        }
                    } else {
                        Logger.warn(`Release action has no holdCode for ${orderId}`);
                        analysis.processError = Messages.ERRORS.HOLD_CODE_MISSING;
                    }
                }

                if (modified) {
                    await SMCApiService.updateOrder(orderData.fullData, pricing);
                    if (createActions.length > 0 || createHoldOnlyActions.length > 0) {
                        orderData.fullData = await SMCApiService.fetchOrderFull(orderId);
                        currentVersion = orderData.fullData?.orderId?.version;
                    }
                }
            }

            for (const analysis of createActions) {
                try {
                    const pricingConfig = DataHelpers.getDetentionPricingConfig(analysis.isPickup);
                    await SMCApiService.addPricingLine(orderId, currentVersion, pricingConfig, analysis.charge);
                    analysis.processed = true; analysis.processedAction = 'created'; analysis.processedAmount = analysis.charge;
                    chargeAdded = true;
                    if (createActions.indexOf(analysis) < createActions.length - 1 || createHoldOnlyActions.length > 0) {
                        orderData.fullData = await SMCApiService.fetchOrderFull(orderId);
                        currentVersion = orderData.fullData?.orderId?.version;
                    }
                } catch (error) { analysis.processed = false; analysis.processError = error.message; Logger.error(`Create failed: ${error.message}`); }
            }

            for (const analysis of createHoldOnlyActions) {
                try {
                    const pricingConfig = DataHelpers.getDetentionPricingConfig(analysis.isPickup);
                    await SMCApiService.addPricingLine(orderId, currentVersion, pricingConfig, 0);
                    analysis.processed = true; analysis.processedAction = 'hold_created'; analysis.processedAmount = 0;
                    Logger.info(`Created $0 hold for ${analysis.isPickup ? 'SHIPPER' : 'RECEIVER'} (Analysis Only)`);
                    if (createHoldOnlyActions.indexOf(analysis) < createHoldOnlyActions.length - 1) {
                        orderData.fullData = await SMCApiService.fetchOrderFull(orderId);
                        currentVersion = orderData.fullData?.orderId?.version;
                    }
                } catch (error) { analysis.processed = false; analysis.processError = error.message; Logger.error(`Create $0 hold failed: ${error.message}`); }
            }

            // Comments: lane mismatch gets TWO comments (timestamp + lane mismatch detail)
            if (chargeAdded) {
                await SMCApiService.addComment(orderId, Messages.COMMENTS.ADD_CHARGE);
            } else if (releaseProcessed) {
                if (hasReleasedWithLaneMismatch) {
                    // Comment A: Timestamp comment
                    await SMCApiService.addComment(orderId, Messages.COMMENTS.BORROWED_TIMESTAMP_RELEASE);
                    // Comment B: Lane mismatch detail comment
                    if (orderData.laneValidation) {
                        const laneMismatchComment = LaneValidator.generateComment(orderData.laneValidation);
                        await SMCApiService.addComment(orderId, laneMismatchComment);
                    }
                } else if (hasReleasedWithBorrowed) {
                    await SMCApiService.addComment(orderId, Messages.COMMENTS.BORROWED_TIMESTAMP_RELEASE);
                } else {
                    await SMCApiService.addComment(orderId, Messages.COMMENTS.RELEASE_HOLD);
                }
            }

            PerformanceMonitor.recordOrderProcessed();
        };

        const processSingleOrder = async (orderId, isBatchMode = false) => {
            const orderData = initializeOrderData(orderId);

            const [viewData, smcExecResult] = await Promise.all([
                HttpClient.request({ method: 'GET', url: `${CONFIG.URLS.SMC_BASE}/shipper/order/view/${encodeURIComponent(orderId)}`, headers: { 'Accept': 'application/json' } }, 'Order View', circuitBreakers.smc),
                ErrorHandler.wrap(() => FMCApiService.fetchSMCExecution(orderId), ErrorType.NETWORK, null, false)
            ]);

            orderData.viewData = viewData; orderData.smcExecutionData = smcExecResult;
            orderData.shipperName = viewData?.shipperDetails?.shipperName || 'Unknown';

            const sowValidation = SOWConfigManager.validateShipper(orderData.shipperName);
            if (!sowValidation.valid) {
                if (!isBatchMode) throw ErrorHandler.create(ErrorType.SOW, sowValidation.error, { orderId, shipper: orderData.shipperName });
                AppState.addBatchReportEntry({ orderId, shipper: orderData.shipperName, action: ActionDisplayConfig.ERROR.reportTerm, amount: '-', status: 'Error', notes: sowValidation.error });
                return null;
            }
            orderData.sowConfig = sowValidation.config;

            if (orderData.smcExecutionData?.tourId) {
                try {
                    orderData.fmcData = await FMCApiService.fetchFMCByTourId(orderData.smcExecutionData.tourId);
                    const matchingVR = FMCApiService.findMatchingVR(orderData.fmcData?.records, orderData.smcExecutionData.contractedLane);
                    orderData.matchingVR = matchingVR;
                    if (matchingVR) {
                        orderData.fmcTimestamps = FMCApiService.extractTimestamps(matchingVR);
                        if (orderData.fmcTimestamps) orderData.fmcTimestamps = FMCApiService.fillMissingTimestampsFromTour(orderData.fmcTimestamps, orderData.fmcData?.records, matchingVR.vehicleRunId);

                        // ★ LANE VALIDATION ★
                        orderData.laneValidation = LaneValidator.validate(viewData, matchingVR);
                    }
                } catch (fmcError) { Logger.warn('FMC fetch failed', fmcError.message); orderData.fmcTimestamps = null; }
            }

            analyzeStops(orderData);

            if (ModeManager.isReviewMode()) {
                for (const analysis of orderData.analysisResults) {
                    if (analysis?.action === ActionType.PENDING_APPROVAL) analysis.approvalNote = Messages.INFO.APPROVAL_REQUIRED_NOTE;
                }
                addToBatchReport(orderData);
                CacheManager.invalidate(orderId);
                return orderData;
            }

            const pendingApprovals = orderData.analysisResults.filter(r => r?.action === ActionType.PENDING_APPROVAL);

            if (pendingApprovals.length > 0 && isBatchMode) {
                AppState.addPendingApprovalOrder(orderData);
                Logger.debug(`Order ${orderId} queued for approval`);
                return orderData;
            }

            const actionsNeeded = orderData.analysisResults.filter(r =>
                r?.action === ActionType.ADD_CHARGE_UPDATE || r?.action === ActionType.ADD_CHARGE_CREATE ||
                r?.action === ActionType.RELEASE_HOLD || r?.action === ActionType.CREATE_HOLD_ONLY
            );

            if (actionsNeeded.length > 0) {
                orderData.fullData = await SMCApiService.fetchOrderFull(orderId);
                await executeActionsWithRetry(orderData);
            }

            addToBatchReport(orderData);
            CacheManager.invalidate(orderId);
            return orderData;
        };

        const processOrderWithRetry = async (orderId, attempt = 1) => {
            try { return await processSingleOrder(orderId, true); }
            catch (error) {
                if (ErrorHandler.isRateLimitError(error)) { Logger.warn(`Rate limited on ${orderId}, waiting...`); await sleep(CONFIG.BATCH.RATE_LIMIT_DELAY * 3); }
                if (attempt < CONFIG.API.MAX_RETRIES && ErrorHandler.isRetryableError(error)) { Logger.warn(`Retry ${attempt + 1} for ${orderId}`); await sleep(CONFIG.BATCH.RATE_LIMIT_DELAY * attempt); return processOrderWithRetry(orderId, attempt + 1); }
                throw error;
            }
        };

        // Fix #11: use cached progressDOM refs
        const throttledUIUpdate = () => {
            const now = Date.now();
            if (now - lastUIUpdate < CONFIG.BATCH.UI_UPDATE_INTERVAL) return;
            lastUIUpdate = now;
            const processedOrders = AppState.get('processedOrders');
            const failedOrders = AppState.get('failedOrders');
            const totalOrders = AppState.get('totalOrders');
            UIController.updateEnhancedBatchProgress({ processed: processedOrders.size + failedOrders.length, success: processedOrders.size, failed: failedOrders.length, total: totalOrders, startTime: startTime });
        };

        const processChunk = async (chunk, chunkIndex) => {
            const parallelBatches = chunkArray(chunk, CONFIG.BATCH.PARALLEL_SIZE);
            for (const parallelBatch of parallelBatches) {
                if (batchState === BatchState.CANCELLED || batchState === BatchState.PAUSED) break;
                const results = await Promise.allSettled(parallelBatch.map(orderId => processOrderWithRetry(orderId)));
                for (let i = 0; i < results.length; i++) {
                    const orderId = parallelBatch[i]; const result = results[i];
                    if (result.status === 'fulfilled' && result.value) AppState.addProcessedOrder(orderId, result.value);
                    else { const errorMsg = result.reason?.message || 'Unknown error'; AppState.addFailedOrder(orderId, errorMsg); AppState.addBatchReportEntry({ orderId, shipper: 'Unknown', action: ActionDisplayConfig.ERROR.reportTerm, amount: '-', status: 'Error', notes: errorMsg }); }
                    throttledUIUpdate();
                }
                if (batchState === BatchState.RUNNING) await sleep(CONFIG.BATCH.RATE_LIMIT_DELAY);
            }
        };

        // Fix #4: handleApproved uses analysis.holdCode consistently
        const handleApprovalFlow = async (orderData, pendingApprovals) => {
            await ensureHoldsExist(orderData, pendingApprovals);
            const totalCharge = pendingApprovals.reduce((sum, a) => sum + (a?.charge || 0), 0);
            const stopDetails = pendingApprovals.map(a => ({ stopType: a?.isPickup ? 'SHIPPER' : 'RECEIVER', charge: a?.charge || 0 }));
            const decision = await ApprovalPopup.show(orderData, totalCharge, stopDetails);
            await processApprovalDecision(orderData, decision, pendingApprovals);
        };

        const ensureHoldsExist = async (orderData, pendingApprovals) => {
            if (!orderData.fullData) orderData.fullData = await SMCApiService.fetchOrderFull(orderData.orderId);
            const pricing = orderData.fullData?.orderDetails?.shipperPricing?.pricing || [];
            const holdInfo = DataHelpers.findDetentionHolds(pricing);
            for (const analysis of pendingApprovals) {
                const isPickup = analysis?.isPickup;
                const holdExists = isPickup ? holdInfo.shipperExists : holdInfo.receiverExists;
                if (!holdExists) {
                    const pricingConfig = DataHelpers.getDetentionPricingConfig(isPickup);
                    try {
                        await SMCApiService.addPricingLine(orderData.orderId, orderData.fullData?.orderId?.version, pricingConfig, 0);
                        Logger.info(`Created $0 hold for ${isPickup ? 'SHIPPER' : 'RECEIVER'}`);
                        orderData.fullData = await SMCApiService.fetchOrderFull(orderData.orderId);
                        analysis.hasHold = true; analysis.holdCode = pricingConfig.pricingCode;
                    } catch (error) { Logger.error(`Failed to create $0 hold: ${error.message}`); }
                }
            }
        };

        const processApprovalDecision = async (orderData, decision, pendingApprovals) => {
            orderData.fullData = await SMCApiService.fetchOrderFull(orderData.orderId);
            switch (decision.decision) {
                case 'YES': await handleApproved(orderData, decision.authorizationNumber, pendingApprovals); break;
                case 'NO': await handleDeclined(orderData, pendingApprovals); break;
                case 'SKIP': handleSkipped(pendingApprovals); break;
                case 'TIMEOUT': handleTimedOut(pendingApprovals); break;
            }
        };

        // Fix #4: uses analysis.holdCode instead of string heuristic
        const handleApproved = async (orderData, authNumber, pendingApprovals) => {
            const pricing = [...(orderData.fullData?.orderDetails?.shipperPricing?.pricing || [])];
            let modified = false;
            for (const analysis of pendingApprovals) {
                if (!analysis?.holdCode) {
                    Logger.warn(`Approval: No holdCode for ${analysis?.isPickup ? 'SHIPPER' : 'RECEIVER'}`);
                    analysis.processError = Messages.ERRORS.HOLD_CODE_MISSING;
                    continue;
                }
                const itemIndex = pricing.findIndex(p => p?.pricingCode === analysis.holdCode);
                if (itemIndex !== -1) {
                    pricing[itemIndex] = { ...pricing[itemIndex], price: { ...pricing[itemIndex].price, value: analysis?.charge || 0 } };
                    analysis.processed = true; analysis.processedAction = 'updated'; analysis.processedAmount = analysis?.charge || 0; analysis.authorizationNumber = authNumber;
                    modified = true;
                } else {
                    Logger.warn(`Approval: holdCode ${analysis.holdCode} not found in pricing`);
                    analysis.processError = Messages.ERRORS.HOLD_CODE_NOT_FOUND(analysis.holdCode, orderData.orderId);
                }
            }
            if (modified) {
                await SMCApiService.updateOrder(orderData.fullData, pricing);
                const comment = authNumber ? Messages.COMMENTS.CHARGE_WITH_AUTH(authNumber) : Messages.COMMENTS.ADD_CHARGE;
                await SMCApiService.addComment(orderData.orderId, comment);
            }
        };

        // Fix #4: uses analysis.holdCode instead of string heuristic
        const handleDeclined = async (orderData, pendingApprovals) => {
            const pricing = [...(orderData.fullData?.orderDetails?.shipperPricing?.pricing || [])];
            let modified = false;
            let hasDeclinedWithBorrowed = false;

            for (const analysis of pendingApprovals) {
                if (!analysis?.holdCode) {
                    Logger.warn(`Decline: No holdCode for ${analysis?.isPickup ? 'SHIPPER' : 'RECEIVER'}`);
                    analysis.processError = Messages.ERRORS.HOLD_CODE_MISSING;
                    continue;
                }
                const itemIndex = pricing.findIndex(p => p?.pricingCode === analysis.holdCode);
                if (itemIndex !== -1) {
                    pricing.splice(itemIndex, 1);
                    analysis.processed = true; analysis.processedAction = 'released';
                    modified = true;
                    if (analysis.hasBorrowedTimestamps) hasDeclinedWithBorrowed = true;
                } else {
                    Logger.warn(`Decline: holdCode ${analysis.holdCode} not found in pricing`);
                    analysis.processError = Messages.ERRORS.HOLD_CODE_NOT_FOUND(analysis.holdCode, orderData.orderId);
                }
            }

            if (modified) {
                await SMCApiService.updateOrder(orderData.fullData, pricing);
                if (hasDeclinedWithBorrowed) {
                    await SMCApiService.addComment(orderData.orderId, Messages.COMMENTS.BORROWED_TIMESTAMP_RELEASE);
                } else {
                    await SMCApiService.addComment(orderData.orderId, Messages.COMMENTS.APPROVAL_DECLINED);
                }
            }
        };

        const handleSkipped = (pendingApprovals) => { for (const analysis of pendingApprovals) { if (analysis) { analysis.processed = true; analysis.processedAction = 'skipped'; } } };
        const handleTimedOut = (pendingApprovals) => { for (const analysis of pendingApprovals) { if (analysis) { analysis.processed = true; analysis.processedAction = 'timeout'; } } };

        const processPendingApprovalOrders = async () => {
            const pendingOrders = AppState.get('pendingApprovalOrders');
            if (pendingOrders.length === 0) return;
            UIController.updateBatchStatus(`Processing ${pendingOrders.length} orders requiring approval...`);
            for (let i = 0; i < pendingOrders.length; i++) {
                if (batchState === BatchState.CANCELLED) break;
                const pendingOrder = pendingOrders[i];
                UIController.updateBatchStatus(`Approval ${i + 1}/${pendingOrders.length}: ${pendingOrder?.orderId || 'Unknown'}`);
                await processApproval(pendingOrder);
                if (i < pendingOrders.length - 1) await sleep(CONFIG.BATCH.PAUSE_CHECK_INTERVAL);
            }
        };

        const processApproval = async (orderData) => {
            const pendingApprovals = (orderData?.analysisResults || []).filter(r => r?.action === ActionType.PENDING_APPROVAL);
            if (pendingApprovals.length === 0) { addToBatchReport(orderData); return; }
            await ensureHoldsExist(orderData, pendingApprovals);
            const totalCharge = pendingApprovals.reduce((sum, a) => sum + (a?.charge || 0), 0);
            const stopDetails = pendingApprovals.map(a => ({ stopType: a?.isPickup ? 'SHIPPER' : 'RECEIVER', charge: a?.charge || 0 }));
            const decision = await ApprovalPopup.show(orderData, totalCharge, stopDetails);
            await processApprovalDecision(orderData, decision, pendingApprovals);
            addToBatchReport(orderData);
            CacheManager.invalidate(orderData?.orderId);
        };

        const saveProgress = (orderIds, chunkIndex) => {
            const processedOrders = AppState.get('processedOrders');
            const failedOrders = AppState.get('failedOrders');
            const batchReportData = AppState.get('batchReportData');
            ProgressManager.save({ orderIds, chunkIndex, processedCount: processedOrders.size, failedCount: failedOrders.length, batchReportData, processedIds: Array.from(processedOrders.keys()) });
        };

        const initializeBatch = (orderIds) => {
            const chunks = chunkArray(orderIds, CONFIG.BATCH.CHUNK_SIZE);
            AppState.resetBatch();
            AppState.update({ currentOrderIds: orderIds, currentIndex: 0, totalOrders: orderIds.length, isProcessing: true, isSingleMode: orderIds.length === 1, batchState: BatchState.RUNNING, batchStartTime: Date.now(), currentChunk: 0, totalChunks: chunks.length, processedOrders: new Map(), failedOrders: [] });
            batchState = BatchState.RUNNING; startTime = Date.now(); progressDOM = null;
            Logger.info(`Batch initialized: ${orderIds.length} orders in ${chunks.length} chunks (Mode: ${ModeManager.getMode()})`);
            Telemetry.track(TelemetryEventType.BATCH_START, { orderCount: orderIds.length, chunkCount: chunks.length, mode: ModeManager.getMode() });
        };

        const finalizeBatch = () => {
            TokenManager.stopAutoRefresh(); ProgressManager.clear(); progressDOM = null;
            const processedOrders = AppState.get('processedOrders');
            const failedOrders = AppState.get('failedOrders');
            const elapsed = startTime ? Date.now() - startTime : 0;
            Logger.info(`Batch complete: ${processedOrders.size} success, ${failedOrders.length} failed in ${Helpers.formatElapsed(elapsed)} (Mode: ${ModeManager.getMode()})`);
            Telemetry.track(TelemetryEventType.BATCH_COMPLETE, { successCount: processedOrders.size, failedCount: failedOrders.length, duration: elapsed, mode: ModeManager.getMode() });
            AppState.update({ isProcessing: false, batchState: BatchState.COMPLETED });
            batchState = BatchState.COMPLETED;
            if (!AppState.get('isSingleMode')) {
                UIController.showBatchComplete(AppState.get('batchReportData'));
                UIController.showToast(Messages.SUCCESS.BATCH_COMPLETE(processedOrders.size, failedOrders.length), failedOrders.length > 0 ? 'warning' : 'success');
            }
        };

        const processEnhancedBatch = async (orderIds) => {
            const chunks = chunkArray(orderIds, CONFIG.BATCH.CHUNK_SIZE);
            UIController.showEnhancedBatchProgress(orderIds.length, chunks.length);
            for (let chunkIndex = 0; chunkIndex < chunks.length; chunkIndex++) {
                if (batchState === BatchState.CANCELLED) { Logger.info(Messages.INFO.BATCH_CANCELLED); break; }
                while (batchState === BatchState.PAUSED) { await sleep(CONFIG.BATCH.PAUSE_CHECK_INTERVAL); if (batchState === BatchState.CANCELLED) break; }
                AppState.set('currentChunk', chunkIndex);
                const chunk = chunks[chunkIndex];
                await ensureValidToken();
                UIController.updateBatchStatus(`Processing chunk ${chunkIndex + 1}/${chunks.length}...`);
                await processChunk(chunk, chunkIndex);
                saveProgress(orderIds, chunkIndex);
                if (chunkIndex < chunks.length - 1 && batchState === BatchState.RUNNING) { UIController.updateBatchStatus(Messages.INFO.COOLING_DOWN); await sleep(CONFIG.BATCH.CHUNK_DELAY); }
            }
            if (batchState !== BatchState.CANCELLED && ModeManager.isActionMode()) await processPendingApprovalOrders();
        };

        const processSingleOrderDetailed = async (orderId) => {
            const isReview = ModeManager.isReviewMode();
            const steps = [
                { id: 'token', text: 'Checking authentication...', icon: '🔐' },
                { id: 'fetch', text: 'Fetching order & tour data...', icon: '📥' },
                { id: 'fmc', text: 'Fetching FMC timestamps...', icon: '⏱️' },
                { id: 'lane', text: 'Validating contracted lane...', icon: '🛤️' },
                { id: 'analyze', text: 'Analyzing detention...', icon: '🔍' },
                { id: 'process', text: isReview ? 'Generating review...' : 'Processing actions...', icon: isReview ? '📋' : '⚡' },
                { id: 'complete', text: 'Complete!', icon: '✅' }
            ];
            UIController.showProgress(steps);
            const orderData = initializeOrderData(orderId);

            try {
                UIController.updateProgressStep('token', 'active');
                await sleep(100);
                UIController.updateProgressStep('token', 'completed', 'Token ready');

                UIController.updateProgressStep('fetch', 'active');
                const [viewData, smcExecResult] = await Promise.all([
                    HttpClient.request({ method: 'GET', url: `${CONFIG.URLS.SMC_BASE}/shipper/order/view/${encodeURIComponent(orderId)}`, headers: { 'Accept': 'application/json' } }, 'Order View', circuitBreakers.smc),
                    ErrorHandler.wrap(() => FMCApiService.fetchSMCExecution(orderId), ErrorType.NETWORK, null, false)
                ]);
                orderData.viewData = viewData; orderData.smcExecutionData = smcExecResult;
                orderData.shipperName = viewData?.shipperDetails?.shipperName || 'Unknown';
                const sowValidation = SOWConfigManager.validateShipper(orderData.shipperName);
                if (!sowValidation.valid) throw ErrorHandler.create(ErrorType.SOW, sowValidation.error, { orderId, shipper: orderData.shipperName });
                orderData.sowConfig = sowValidation.config;
                const tourInfo = smcExecResult ? `Tour: ${String(smcExecResult.tourId || '').substring(0, 15)}...` : 'Tour unavailable';
                UIController.updateProgressStep('fetch', 'completed', tourInfo);

                UIController.updateProgressStep('fmc', 'active');
                if (orderData.smcExecutionData?.tourId) {
                    try {
                        orderData.fmcData = await FMCApiService.fetchFMCByTourId(orderData.smcExecutionData.tourId);
                        const matchingVR = FMCApiService.findMatchingVR(orderData.fmcData?.records, orderData.smcExecutionData.contractedLane);
                        orderData.matchingVR = matchingVR;
                        if (matchingVR) {
                            orderData.fmcTimestamps = FMCApiService.extractTimestamps(matchingVR);
                            if (orderData.fmcTimestamps) orderData.fmcTimestamps = FMCApiService.fillMissingTimestampsFromTour(orderData.fmcTimestamps, orderData.fmcData?.records, matchingVR.vehicleRunId);
                        }
                        UIController.updateProgressStep('fmc', 'completed', 'Timestamps loaded');
                    } catch (fmcError) { Logger.warn('FMC fetch failed', fmcError.message); UIController.updateProgressStep('fmc', 'error', 'FMC unavailable'); }
                } else { UIController.updateProgressStep('fmc', 'error', 'No Tour ID'); }

                // ★ LANE VALIDATION STEP ★
                UIController.updateProgressStep('lane', 'active');
                if (orderData.matchingVR) {
                    orderData.laneValidation = LaneValidator.validate(viewData, orderData.matchingVR);
                    if (orderData.laneValidation.isMatch) {
                        UIController.updateProgressStep('lane', 'completed', '✓ Lane match confirmed');
                    } else {
                        UIController.updateProgressStep('lane', 'error', '🚫 Lane mismatch detected');
                    }
                } else {
                    UIController.updateProgressStep('lane', 'completed', 'No VR — skipped');
                }

                UIController.updateProgressStep('analyze', 'active');
                analyzeStops(orderData);
                const hasLaneMismatch = orderData.analysisResults.some(a => a?.laneMismatch);
                if (hasLaneMismatch) {
                    UIController.updateProgressStep('analyze', 'completed', '🚫 Lane mismatch — no charge');
                } else {
                    UIController.updateProgressStep('analyze', 'completed', `${orderData.analysisResults.length} stops analyzed`);
                }

                if (isReview) {
                    for (const analysis of orderData.analysisResults) { if (analysis?.action === ActionType.PENDING_APPROVAL) analysis.approvalNote = Messages.INFO.APPROVAL_REQUIRED_NOTE; }
                    if (hasLaneMismatch) {
                        const hasHolds = orderData.analysisResults.some(a => a?.hasHold);
                        UIController.updateProgressStep('process', 'completed', hasHolds ? 'Would release holds (Lane Mismatch)' : 'No charge (Lane Mismatch)');
                    } else {
                        UIController.updateProgressStep('process', 'completed', Messages.INFO.REVIEW_MODE_NO_ACTIONS);
                    }
                    UIController.updateProgressStep('complete', 'completed');
                    CacheManager.invalidate(orderId);
                    AppState.set('singleOrderData', orderData);
                    UIController.displaySingleOrderResults(orderData);
                    return;
                }

                const pendingApprovals = orderData.analysisResults.filter(r => r?.action === ActionType.PENDING_APPROVAL);
                if (pendingApprovals.length > 0) {
                    UIController.updateProgressStep('process', 'active', 'Approval required...');
                    await handleApprovalFlow(orderData, pendingApprovals);
                    UIController.updateProgressStep('process', 'completed', 'Processed');
                } else {
                    const actionsNeeded = orderData.analysisResults.filter(r => r?.action === ActionType.ADD_CHARGE_UPDATE || r?.action === ActionType.ADD_CHARGE_CREATE || r?.action === ActionType.RELEASE_HOLD || r?.action === ActionType.CREATE_HOLD_ONLY);
                    const analysisOnlyActions = orderData.analysisResults.filter(r => r?.action === ActionType.ANALYSIS_ONLY);
                    if (actionsNeeded.length === 0) {
                        if (analysisOnlyActions.length > 0) UIController.updateProgressStep('process', 'completed', 'Analysis only (no auto-charge)');
                        else UIController.updateProgressStep('process', 'completed', 'No actions needed');
                    } else {
                        const releaseCount = actionsNeeded.filter(a => a?.action === ActionType.RELEASE_HOLD).length;
                        const isLaneMismatchRelease = actionsNeeded.some(a => a?.laneMismatch);
                        UIController.updateProgressStep('process', 'active', isLaneMismatchRelease ? 'Releasing holds (Lane Mismatch)...' : `Processing ${actionsNeeded.length} action(s)...`);
                        orderData.fullData = await SMCApiService.fetchOrderFull(orderId);
                        await executeActionsWithRetry(orderData);
                        UIController.updateProgressStep('process', 'completed', isLaneMismatchRelease ? `${releaseCount} hold(s) released (Lane Mismatch)` : `${actionsNeeded.length} action(s) completed`);
                    }
                }

                UIController.updateProgressStep('complete', 'completed');
                CacheManager.invalidate(orderId);
                AppState.set('singleOrderData', orderData);
                UIController.displaySingleOrderResults(orderData);
            } catch (error) { Logger.error('Processing error', error.message); UIController.showProcessingError(error.message); }
        };

        const promptResumeProgress = (savedProgress) => {
            return new Promise(resolve => {
                const overlay = document.createElement('div'); overlay.id = 'd-dart-resume-overlay';
                overlay.innerHTML = `<div class="d-dart-resume-popup"><div class="d-dart-resume-title">📋 Resume Previous Batch?</div><div class="d-dart-resume-info">Found saved progress:<br><strong>${savedProgress.processedCount}</strong> orders processed<br><strong>${savedProgress.failedCount}</strong> failed<br><strong>${savedProgress.orderIds.length - savedProgress.processedCount - savedProgress.failedCount}</strong> remaining</div><div class="d-dart-resume-buttons"><button class="d-dart-resume-btn yes" id="d-dart-resume-yes">✅ Resume</button><button class="d-dart-resume-btn no" id="d-dart-resume-no">🔄 Start Fresh</button></div></div>`;
                document.body.appendChild(overlay);
                document.getElementById('d-dart-resume-yes').onclick = () => { overlay.remove(); resolve(true); };
                document.getElementById('d-dart-resume-no').onclick = () => { overlay.remove(); ProgressManager.clear(); resolve(false); };
            });
        };

        const resumeFromProgress = async (savedProgress, orderIds) => {
            Logger.info(`Resuming batch: ${savedProgress.processedCount} already processed`);
            AppState.set('batchReportData', savedProgress.batchReportData || []);
            const processedIds = new Set(savedProgress.processedIds || []);
            const remainingOrders = orderIds.filter(id => !processedIds.has(id));
            if (remainingOrders.length === 0) { UIController.showBatchComplete(savedProgress.batchReportData); return; }
            await startFreshBatch(remainingOrders);
        };

        const startFreshBatch = async (orderIds) => {
            initializeBatch(orderIds);
            if (ModeManager.isActionMode()) {
                const tokenOk = await TokenManager.ensure();
                if (!tokenOk) { UIController.showProcessingError(Messages.ERRORS.TOKEN_MISSING); AppState.set('isProcessing', false); return; }
                if (orderIds.length > 10) TokenManager.startAutoRefresh();
            }
            if (orderIds.length === 1) await processSingleOrderDetailed(orderIds[0]);
            else await processEnhancedBatch(orderIds);
            finalizeBatch();
        };

        const isSameBatch = (savedIds, currentIds) => { if (!savedIds || savedIds.length !== currentIds.length) return false; return savedIds.every((id, i) => id === currentIds[i]); };

        return {
            async processBatch(orderIds) {
                if (!SOWConfigManager.isLoaded()) { UIController.showToast(Messages.ERRORS.SOW_SERVER_UNREACHABLE, 'error'); return; }
                if (orderIds.length > CONFIG.BATCH.MAX_ORDERS_PER_SESSION) { UIController.showToast(Messages.ERRORS.BATCH_TOO_LARGE(CONFIG.BATCH.MAX_ORDERS_PER_SESSION), 'error'); return; }
                const savedProgress = ProgressManager.load();
                if (savedProgress && isSameBatch(savedProgress.orderIds, orderIds)) { const resume = await promptResumeProgress(savedProgress); if (resume) return resumeFromProgress(savedProgress, orderIds); }
                await startFreshBatch(orderIds);
            },
            pause() { if (batchState === BatchState.RUNNING) { batchState = BatchState.PAUSED; AppState.set('batchState', BatchState.PAUSED); Logger.info('Batch paused'); UIController.updateBatchStatus(Messages.INFO.BATCH_PAUSED); } },
            resume() { if (batchState === BatchState.PAUSED) { batchState = BatchState.RUNNING; AppState.set('batchState', BatchState.RUNNING); Logger.info('Batch resumed'); } },
            cancel() { batchState = BatchState.CANCELLED; AppState.set('batchState', BatchState.CANCELLED); Logger.info('Batch cancelled'); },
            getState() { return batchState; },
            // Fix #11: expose progressDOM setter for UIController
            setProgressDOM(refs) { progressDOM = refs; },
            getProgressDOM() { return progressDOM; }
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 29: HEALTH CHECK SYSTEM
     * ═══════════════════════════════════════════════════════════════════════════ */

    const HealthCheck = {
        check() {
            return {
                timestamp: Date.now(), version: CONFIG.VERSION, appState: AppState.getSnapshot(),
                operationMode: ModeManager.getMode(),
                tokenStatus: TokenManager.getStatus(), sowStatus: SOWConfigManager.getStatus(),
                circuitBreakers: { smc: circuitBreakers.smc.getState(), fmc: circuitBreakers.fmc.getState(), execution: circuitBreakers.execution.getState() },
                cache: CacheManager.getStats(), performance: PerformanceMonitor.getMetrics(), telemetry: Telemetry.getMetrics(),
                ui: { isMinimized: AppState.get('isMinimized'), isProcessing: AppState.get('isProcessing'), batchState: AppState.get('batchState'), isSettingsOpen: AppState.get('isSettingsOpen') }
            };
        },
        isHealthy() {
            const status = this.check();
            const circuitBreakerOpen = Object.values(status.circuitBreakers).some(s => s?.state === CircuitBreakerState.OPEN);
            const tokenExpired = status.tokenStatus?.status === 'expired';
            const sowNotLoaded = status.sowStatus?.status !== SOWStatus.LOADED;
            return !circuitBreakerOpen && !tokenExpired && !sowNotLoaded;
        },
        getSummary() {
            const status = this.check(); const issues = [];
            if (status.tokenStatus?.status === 'expired') issues.push('Token expired');
            if (status.tokenStatus?.status === 'missing') issues.push('Token missing');
            if (status.sowStatus?.status !== SOWStatus.LOADED) issues.push(`SOW: ${status.sowStatus?.status}`);
            Object.entries(status.circuitBreakers).forEach(([name, state]) => { if (state?.state === CircuitBreakerState.OPEN) issues.push(`${name.toUpperCase()} circuit open`); });
            if (issues.length === 0) return '✅ System healthy';
            return `⚠️ Issues: ${issues.join(', ')}`;
        }
    };

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 30: STYLES (with lane mismatch CSS)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const Styles = `
    #d-dart, #d-dart * { box-sizing: border-box !important; font-family: 'Amazon Ember', 'Segoe UI', Tahoma, sans-serif !important; }
    #d-dart { position: fixed !important; top: ${CONFIG.INITIAL_POSITION.top} !important; left: ${CONFIG.INITIAL_POSITION.left} !important; right: ${CONFIG.INITIAL_POSITION.right} !important; width: ${CONFIG.UI.PANEL_WIDTH}px !important; background: #232F3E !important; border: 2px solid #FF9900 !important; border-radius: 12px !important; box-shadow: 0 8px 32px rgba(0,0,0,0.4) !important; z-index: 2147483647 !important; color: #FFF !important; overflow: hidden !important; transition: box-shadow 0.2s ease, width 0.3s ease, height 0.3s ease, border-radius 0.3s ease !important; }
    #d-dart.dragging { opacity: 0.9 !important; box-shadow: 0 12px 40px rgba(0,0,0,0.6) !important; cursor: grabbing !important; }
    #d-dart.minimized { width: ${CONFIG.UI.PANEL_MIN_WIDTH}px !important; height: ${CONFIG.UI.PANEL_MIN_WIDTH}px !important; border-radius: 50% !important; cursor: grab !important; display: flex !important; align-items: center !important; justify-content: center !important; padding: 0 !important; }
    #d-dart.minimized:hover { box-shadow: 0 0 20px rgba(255,153,0,0.6) !important; transform: scale(1.05) !important; }
    #d-dart.minimized #d-dart-header, #d-dart.minimized #d-dart-body { display: none !important; }
    #d-dart.minimized #d-dart-minimized-icon { display: flex !important; }
    #d-dart-minimized-icon { display: none !important; font-size: 26px !important; align-items: center !important; justify-content: center !important; width: 100% !important; height: 100% !important; cursor: pointer !important; }
    #d-dart.minimized.healthy { box-shadow: 0 0 20px rgba(0,255,136,0.5) !important; }
    #d-dart.minimized.unhealthy { box-shadow: 0 0 20px rgba(255,107,107,0.5) !important; }
    #d-dart-header { background: linear-gradient(90deg, #FF9900 0%, #E88B00 100%) !important; padding: 0 !important; cursor: grab !important; user-select: none !important; display: flex !important; flex-direction: column !important; }
    #d-dart-header:active { cursor: grabbing !important; }
    .d-dart-header-row-1 { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 4px 10px 1px !important; }
    .d-dart-header-row-2 { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 0px 10px 4px !important; }
    .d-dart-app-title { margin: 0 !important; font-size: 12px !important; font-weight: 700 !important; color: #232F3E !important; display: flex !important; align-items: center !important; gap: 4px !important; }
    .d-dart-header-buttons { display: flex !important; gap: 6px !important; }
    .d-dart-header-btn { background: transparent !important; border: none !important; color: #232F3E !important; font-size: 13px !important; cursor: pointer !important; padding: 1px 4px !important; transition: transform 0.2s ease !important; font-weight: bold !important; line-height: 1 !important; }
    .d-dart-header-btn:hover { transform: scale(1.15) !important; }
    .d-dart-header-btn:focus { outline: 2px solid #232F3E !important; outline-offset: 1px !important; }
    .d-dart-mode-toggle { display: flex !important; background: rgba(0,0,0,0.3) !important; border-radius: 4px !important; padding: 1px !important; cursor: pointer !important; gap: 1px !important; }
    .d-dart-mode-toggle:focus { outline: 2px solid #232F3E !important; outline-offset: 1px !important; }
    .d-dart-mode-option { font-size: 8px !important; font-weight: 700 !important; padding: 2px 6px !important; border-radius: 3px !important; transition: all 0.2s ease !important; color: rgba(35,47,62,0.5) !important; user-select: none !important; }
    .d-dart-mode-option.active.action { background: #ff6b6b !important; color: #FFF !important; text-shadow: 0 0 8px rgba(255,107,107,0.5) !important; }
    .d-dart-mode-option.active.review { background: #4dabf7 !important; color: #FFF !important; text-shadow: 0 0 8px rgba(77,171,247,0.5) !important; }
    .d-dart-signature-block { display: flex !important; align-items: baseline !important; gap: 3px !important; white-space: nowrap !important; }
    .d-dart-made-by { font-size: 8px !important; font-weight: 400 !important; color: rgba(35,47,62,0.5) !important; }
    .d-dart-signature { font-size: 10px !important; color: #232F3E !important; font-weight: 700 !important; white-space: nowrap !important; }
    #d-dart-body { padding: 12px !important; max-height: 75vh !important; overflow-y: auto !important; background: #1a242f !important; }
    .d-dart-input-group { display: flex !important; gap: 8px !important; margin-bottom: 12px !important; }
    .d-dart-input { flex: 1 !important; padding: 10px 12px !important; border: 2px solid #37475A !important; border-radius: 6px !important; background: #232F3E !important; color: #FFF !important; font-size: 13px !important; outline: none !important; transition: border-color 0.2s ease !important; }
    .d-dart-input:focus { border-color: #FF9900 !important; }
    .d-dart-input.error { border-color: #ff6b6b !important; animation: d-dart-shake 0.3s !important; }
    .d-dart-input::placeholder { color: #666 !important; font-size: 11px !important; }
    .d-dart-input:disabled { opacity: 0.6 !important; cursor: not-allowed !important; }
    @keyframes d-dart-shake { 0%, 100% { transform: translateX(0); } 25% { transform: translateX(-5px); } 75% { transform: translateX(5px); } }
    @keyframes d-dart-pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.6; } }
    @keyframes d-dart-spin { to { transform: rotate(360deg); } }
    @keyframes d-dart-fadeIn { from { opacity: 0; } to { opacity: 1; } }
    @keyframes d-dart-popupSlide { from { transform: translateY(-20px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
    @keyframes d-dart-pop { from { transform: translate(-50%, -50%) scale(0.5); opacity: 0; } to { transform: translate(-50%, -50%) scale(1); opacity: 1; } }
    .d-dart-btn { padding: 10px 20px !important; background: #FF9900 !important; border: none !important; border-radius: 6px !important; color: #232F3E !important; font-weight: 700 !important; font-size: 13px !important; cursor: pointer !important; transition: all 0.2s ease !important; min-width: 90px !important; position: relative !important; }
    .d-dart-btn:hover:not(:disabled) { background: #FEBD69 !important; transform: translateY(-1px) !important; box-shadow: 0 4px 12px rgba(255,153,0,0.4) !important; }
    .d-dart-btn:focus { outline: 2px solid #FFF !important; outline-offset: 2px !important; }
    .d-dart-btn:disabled { background: #555 !important; cursor: not-allowed !important; }
    .d-dart-btn.loading { color: transparent !important; }
    .d-dart-btn.loading::after { content: ''; position: absolute; width: 16px; height: 16px; top: 50%; left: 50%; margin: -8px 0 0 -8px; border: 2px solid #232F3E; border-top-color: transparent; border-radius: 50%; animation: d-dart-spin 0.8s linear infinite; }
    .d-dart-toast { position: fixed !important; bottom: 25px !important; left: 50% !important; transform: translateX(-50%) translateY(20px) !important; background: #FF9900 !important; color: #232F3E !important; padding: 10px 20px !important; border-radius: 8px !important; font-size: 13px !important; font-weight: 600 !important; z-index: 2147483648 !important; opacity: 0 !important; transition: all 0.3s ease !important; pointer-events: none !important; max-width: 400px !important; text-align: center !important; }
    .d-dart-toast.show { opacity: 1 !important; transform: translateX(-50%) translateY(0) !important; }
    .d-dart-toast.error { background: #ff6b6b !important; color: #FFF !important; }
    .d-dart-toast.success { background: #00ff88 !important; color: #232F3E !important; }
    .d-dart-toast.warning { background: #ffd700 !important; color: #232F3E !important; }
    .d-dart-visually-hidden { position: absolute !important; width: 1px !important; height: 1px !important; padding: 0 !important; margin: -1px !important; overflow: hidden !important; clip: rect(0, 0, 0, 0) !important; white-space: nowrap !important; border: 0 !important; }
    .d-dart-review-banner { background: rgba(77,171,247,0.15) !important; border: 1px solid #4dabf7 !important; border-radius: 6px !important; padding: 10px !important; text-align: center !important; color: #4dabf7 !important; font-size: 12px !important; font-weight: 600 !important; margin-bottom: 10px !important; }
    .d-dart-borrowed-warning { background: rgba(255,212,59,0.15) !important; border: 1px solid #ffd43b !important; border-radius: 4px !important; padding: 6px 10px !important; font-size: 11px !important; color: #ffd43b !important; font-weight: 600 !important; margin-top: 6px !important; text-align: center !important; }
    .d-dart-mismatch-box { border-radius: 6px !important; padding: 10px 12px !important; margin: 8px 12px !important; }
    .d-dart-mismatch-box.under { background: rgba(255,165,0,0.15) !important; border: 1px solid #ffa500 !important; }
    .d-dart-mismatch-box.over { background: rgba(255,107,107,0.15) !important; border: 1px solid #ff6b6b !important; }
    .d-dart-mismatch-box.should-not-exist { background: rgba(255,59,59,0.2) !important; border: 1px solid #ff3b3b !important; }
    .d-dart-mismatch-title { font-size: 12px !important; font-weight: 700 !important; margin-bottom: 8px !important; }
    .d-dart-mismatch-box.under .d-dart-mismatch-title { color: #ffa500 !important; }
    .d-dart-mismatch-box.over .d-dart-mismatch-title { color: #ff6b6b !important; }
    .d-dart-mismatch-box.should-not-exist .d-dart-mismatch-title { color: #ff3b3b !important; }
    .d-dart-mismatch-row { display: flex !important; justify-content: space-between !important; padding: 4px 0 !important; font-size: 12px !important; color: #ccc !important; }
    .d-dart-mismatch-row.diff { font-weight: 700 !important; border-top: 1px solid rgba(255,255,255,0.1) !important; padding-top: 6px !important; margin-top: 4px !important; }
    .d-dart-mismatch-box.under .d-dart-mismatch-row.diff { color: #ffa500 !important; }
    .d-dart-mismatch-box.over .d-dart-mismatch-row.diff { color: #ff6b6b !important; }
    .d-dart-mismatch-box.should-not-exist .d-dart-mismatch-row.diff { color: #ff3b3b !important; }
    .d-dart-banner-value.mismatch-under { color: #ffa500 !important; }
    .d-dart-banner-value.mismatch-over { color: #ff6b6b !important; }
    .d-dart-banner-value.mismatch-should-not-exist { color: #ff3b3b !important; }
    .d-dart-banner-value.charge-correct { color: #00ff88 !important; }
    .d-dart-batch-stat.mismatch-stat { color: #ffa500 !important; }
    .d-dart-batch-stat.correct-stat { color: #00ff88 !important; }
    .d-dart-batch-table tr.mismatch td { background: rgba(255,165,0,0.08) !important; }
    .d-dart-batch-table tr.correct td { background: rgba(0,255,136,0.05) !important; }

    /* Lane Mismatch CSS */
    .d-dart-lane-mismatch-banner { background: rgba(255,59,59,0.12) !important; border: 2px solid #ff6b6b !important; border-radius: 8px !important; padding: 15px !important; margin-bottom: 10px !important; text-align: center !important; }
    .d-dart-lane-mismatch-icon { font-size: 32px !important; margin-bottom: 6px !important; }
    .d-dart-lane-mismatch-title { font-size: 14px !important; font-weight: 700 !important; color: #ff6b6b !important; margin-bottom: 10px !important; text-transform: uppercase !important; letter-spacing: 1px !important; }
    .d-dart-lane-mismatch-detail { background: #1a242f !important; border-radius: 6px !important; padding: 10px 12px !important; margin-bottom: 10px !important; }
    .d-dart-lane-compare-row { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 6px 0 !important; font-size: 12px !important; border-bottom: 1px solid #37475A !important; }
    .d-dart-lane-compare-row:last-child { border-bottom: none !important; }
    .d-dart-lane-compare-label { color: #888 !important; font-weight: 600 !important; }
    .d-dart-lane-compare-value { color: #FFF !important; font-weight: 600 !important; font-family: monospace !important; font-size: 11px !important; }
    .d-dart-lane-mismatch-highlight { color: #ff6b6b !important; background: rgba(255,107,107,0.2) !important; padding: 2px 6px !important; border-radius: 3px !important; font-weight: 700 !important; }
    .d-dart-lane-mismatch-action { font-size: 11px !important; color: #ff9999 !important; font-style: italic !important; }
    .d-dart-banner-value.lane-mismatch { color: #ff6b6b !important; }
    .d-dart-stop-card.d-dart-stop-lane-mismatch { border-left-color: #ff6b6b !important; opacity: 0.8 !important; }
    .d-dart-lane-mismatch-stop-indicator { background: rgba(255,107,107,0.15) !important; border: 1px solid #ff6b6b !important; border-radius: 4px !important; padding: 6px 10px !important; font-size: 11px !important; color: #ff6b6b !important; font-weight: 600 !important; margin-bottom: 8px !important; text-align: center !important; }
    .d-dart-stop-badge.lane-mismatch { background: rgba(255,107,107,0.25) !important; color: #ff6b6b !important; }
    .d-dart-batch-stat.lane-mismatch-stat { color: #ff6b6b !important; }
    .d-dart-batch-table tr.lane-mismatch-row td { background: rgba(255,107,107,0.08) !important; }

    /* SOW Refresh Loading States */
    .d-dart-refresh-sow-btn.loading { position: relative !important; color: transparent !important; pointer-events: none !important; }
    .d-dart-refresh-sow-btn.loading::after { content: '' !important; position: absolute !important; width: 18px !important; height: 18px !important; top: 50% !important; left: calc(50% - 60px) !important; margin: -9px 0 0 0 !important; border: 3px solid rgba(35,47,62,0.3) !important; border-top-color: #232F3E !important; border-radius: 50% !important; animation: d-dart-spin 0.7s linear infinite !important; }
    .d-dart-refresh-sow-btn.loading::before { content: '⏳ Refreshing SOW...' !important; position: absolute !important; top: 50% !important; left: 50% !important; transform: translate(-40%, -50%) !important; color: #232F3E !important; font-size: 13px !important; font-weight: 700 !important; white-space: nowrap !important; }
    .d-dart-refresh-elapsed { text-align: center !important; font-size: 11px !important; color: #888 !important; margin-top: 6px !important; min-height: 16px !important; transition: opacity 0.3s ease !important; }
    .d-dart-shippers-loading { padding: 30px 20px !important; text-align: center !important; }
    .d-dart-loading-skeleton { background: linear-gradient(90deg, #37475A 25%, #485769 50%, #37475A 75%) !important; background-size: 200% 100% !important; animation: d-dart-skeleton-shimmer 1.5s ease-in-out infinite !important; border-radius: 6px !important; height: 48px !important; margin-bottom: 8px !important; }
    @keyframes d-dart-skeleton-shimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }
    .d-dart-loading-text { color: #888 !important; font-size: 12px !important; margin-top: 12px !important; animation: d-dart-pulse 1.2s ease-in-out infinite !important; }
    .d-dart-settings-section { background: #37475A !important; border-radius: 8px !important; padding: 12px !important; margin-bottom: 12px !important; }
    .d-dart-settings-section-title { font-size: 12px !important; font-weight: 700 !important; color: #FF9900 !important; text-transform: uppercase !important; margin-bottom: 10px !important; letter-spacing: 0.5px !important; }
    .d-dart-settings-refresh-section { padding: 12px !important; margin-bottom: 12px !important; }
    .d-dart-refresh-sow-btn { width: 100% !important; padding: 12px 20px !important; background: #FF9900 !important; border: none !important; border-radius: 8px !important; color: #232F3E !important; font-size: 14px !important; font-weight: 700 !important; cursor: pointer !important; transition: all 0.2s ease !important; }
    .d-dart-refresh-sow-btn:hover { background: #FEBD69 !important; transform: translateY(-1px) !important; box-shadow: 0 4px 12px rgba(255,153,0,0.4) !important; }
    .d-dart-refresh-sow-btn:disabled { background: #555 !important; cursor: not-allowed !important; }
    .d-dart-shippers-list { max-height: 350px !important; overflow-y: auto !important; }
    .d-dart-no-results { text-align: center !important; padding: 30px !important; color: #888 !important; font-size: 13px !important; }
    .d-dart-shipper-card-settings { background: #1a242f !important; border-radius: 6px !important; margin-bottom: 8px !important; border-left: 3px solid #37475A !important; overflow: hidden !important; }
    .d-dart-shipper-card-settings.status-active { border-left-color: #00ff88 !important; }
    .d-dart-shipper-card-settings.status-inactive { border-left-color: #888 !important; }
    .d-dart-shipper-card-settings.status-error { border-left-color: #ff6b6b !important; }
    .d-dart-shipper-header-settings { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 10px 12px !important; cursor: pointer !important; transition: background 0.2s ease !important; }
    .d-dart-shipper-header-settings:hover { background: rgba(255,153,0,0.1) !important; }
    .d-dart-shipper-info { display: flex !important; align-items: center !important; gap: 8px !important; }
    .d-dart-shipper-status-icon { font-size: 14px !important; }
    .d-dart-shipper-name-settings { font-size: 13px !important; font-weight: 600 !important; color: #FFF !important; }
    .d-dart-shipper-summary { display: flex !important; align-items: center !important; gap: 10px !important; }
    .d-dart-shipper-rate, .d-dart-shipper-max { font-size: 11px !important; color: #888 !important; }
    .d-dart-validation-error-badge { font-size: 10px !important; color: #ff6b6b !important; background: rgba(255,107,107,0.15) !important; padding: 2px 8px !important; border-radius: 4px !important; }
    .d-dart-expand-btn { background: transparent !important; border: none !important; color: #888 !important; font-size: 12px !important; cursor: pointer !important; padding: 4px 8px !important; transition: color 0.2s ease !important; }
    .d-dart-expand-btn:hover { color: #FF9900 !important; }
    .d-dart-shipper-details { max-height: 0 !important; overflow: hidden !important; transition: max-height 0.3s ease, padding 0.3s ease !important; background: #232F3E !important; }
    .d-dart-shipper-details.expanded { max-height: 600px !important; padding: 12px !important; border-top: 1px solid #37475A !important; }
    .d-dart-validation-errors { background: rgba(255,107,107,0.1) !important; border: 1px solid #ff6b6b !important; border-radius: 6px !important; padding: 10px !important; margin-bottom: 12px !important; }
    .d-dart-error-title { font-size: 12px !important; font-weight: 600 !important; color: #ff6b6b !important; margin-bottom: 6px !important; }
    .d-dart-error-list { margin: 0 !important; padding-left: 20px !important; font-size: 11px !important; color: #ff9999 !important; }
    .d-dart-error-list li { margin-bottom: 4px !important; }
    .d-dart-details-grid { display: grid !important; grid-template-columns: repeat(2, 1fr) !important; gap: 10px !important; }
    .d-dart-details-section { background: #37475A !important; border-radius: 6px !important; padding: 10px !important; }
    .d-dart-details-title { font-size: 10px !important; font-weight: 700 !important; color: #FF9900 !important; text-transform: uppercase !important; margin-bottom: 8px !important; padding-bottom: 4px !important; border-bottom: 1px solid #485769 !important; }
    .d-dart-details-row { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 4px 0 !important; font-size: 11px !important; }
    .d-dart-details-label { color: #888 !important; }
    .d-dart-details-value { color: #FFF !important; font-weight: 500 !important; }
    .d-dart-details-value.yes { color: #00ff88 !important; }
    .d-dart-details-value.no { color: #ff6b6b !important; }
    .d-dart-settings-inline { padding: 0 !important; }
    .d-dart-settings-inline-header { display: flex !important; align-items: center !important; gap: 12px !important; padding: 12px !important; background: linear-gradient(90deg, #FF9900, #E88B00) !important; margin: -12px -12px 12px -12px !important; border-radius: 0 !important; }
    .d-dart-back-btn { background: rgba(0,0,0,0.2) !important; border: none !important; color: #232F3E !important; font-size: 13px !important; font-weight: 700 !important; padding: 6px 12px !important; border-radius: 6px !important; cursor: pointer !important; transition: all 0.2s ease !important; }
    .d-dart-back-btn:hover { background: rgba(0,0,0,0.3) !important; transform: translateX(-2px) !important; }
    .d-dart-back-btn:focus { outline: 2px solid #232F3E !important; outline-offset: 2px !important; }
    .d-dart-settings-inline-title { font-size: 16px !important; font-weight: 700 !important; color: #232F3E !important; }
    .d-dart-detention-banner { background: #232F3E !important; border: 1px solid #FF9900 !important; border-radius: 8px !important; margin-bottom: 10px !important; overflow: hidden !important; }
    .d-dart-banner-header { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 10px 12px !important; background: linear-gradient(90deg, rgba(255,153,0,0.2), rgba(255,153,0,0.05)) !important; border-bottom: 1px solid #37475A !important; }
    .d-dart-banner-title { font-size: 12px !important; font-weight: 700 !important; color: #FF9900 !important; text-transform: uppercase !important; letter-spacing: 0.5px !important; }
    .d-dart-banner-toggle { background: transparent !important; border: 1px solid #FF9900 !important; border-radius: 4px !important; color: #FF9900 !important; font-size: 11px !important; padding: 4px 10px !important; cursor: pointer !important; transition: all 0.2s ease !important; font-weight: 600 !important; }
    .d-dart-banner-toggle:hover { background: #FF9900 !important; color: #232F3E !important; }
    .d-dart-banner-summary { padding: 12px !important; }
    .d-dart-banner-row { display: flex !important; align-items: center !important; padding: 8px 12px !important; background: #37475A !important; border-radius: 6px !important; margin-bottom: 6px !important; }
    .d-dart-banner-row:last-child { margin-bottom: 0 !important; }
    .d-dart-banner-label { font-size: 12px !important; font-weight: 700 !important; color: #888 !important; width: 90px !important; flex-shrink: 0 !important; }
    .d-dart-banner-value { font-size: 13px !important; font-weight: 700 !important; flex: 1 !important; }
    .d-dart-banner-value.charge-added { color: #00ff88 !important; }
    .d-dart-banner-value.hold-released { color: #4dabf7 !important; }
    .d-dart-banner-value.chargeable { color: #ffd700 !important; }
    .d-dart-banner-value.charge-exists { color: #ffd700 !important; }
    .d-dart-banner-value.pending { color: #888 !important; }
    .d-dart-banner-value.no-charge { color: #ff6b6b !important; }
    .d-dart-banner-value.no-action { color: #888 !important; }
    .d-dart-banner-value.analysis-only { color: #4dabf7 !important; }
    .d-dart-banner-value.hold-created { color: #4dabf7 !important; }
    .d-dart-banner-details { max-height: 0 !important; overflow: hidden !important; transition: max-height 0.3s ease, padding 0.3s ease !important; background: #1a242f !important; }
    .d-dart-banner-details.expanded { max-height: 400px !important; padding: 12px !important; border-top: 1px solid #37475A !important; }
    .d-dart-breakdown-section { background: #232F3E !important; border-radius: 6px !important; padding: 10px 12px !important; margin-bottom: 8px !important; border: 1px solid #37475A !important; }
    .d-dart-breakdown-section:last-child { margin-bottom: 0 !important; }
    .d-dart-breakdown-title { font-size: 11px !important; font-weight: 700 !important; color: #FF9900 !important; margin-bottom: 8px !important; text-transform: uppercase !important; }
    .d-dart-breakdown-content { font-size: 11px !important; color: #ccc !important; line-height: 1.6 !important; }
    .d-dart-breakdown-line { padding: 2px 0 !important; }
    .d-dart-shipper-card { background: #37475A !important; border-radius: 8px !important; padding: 10px !important; margin-bottom: 10px !important; border: 1px solid #485769 !important; }
    .d-dart-header-row { display: flex !important; justify-content: space-between !important; align-items: center !important; margin-bottom: 8px !important; padding-bottom: 8px !important; border-bottom: 1px solid #485769 !important; }
    .d-dart-shipper-name { font-size: 14px !important; font-weight: 700 !important; color: #FF9900 !important; }
    .d-dart-header-badges { display: flex !important; align-items: center !important; gap: 8px !important; }
    .d-dart-status-badge { padding: 4px 10px !important; border-radius: 4px !important; font-size: 10px !important; font-weight: 600 !important; color: #FFF !important; text-transform: uppercase !important; }
    .d-dart-sow-badge { font-size: 12px !important; font-weight: 700 !important; cursor: pointer !important; padding: 2px 6px !important; border-radius: 4px !important; color: #FF9900 !important; text-shadow: 0 0 5px rgba(255, 153, 0, 0.5) !important; animation: d-dart-glow-pulse 2s ease-in-out infinite !important; }
    .d-dart-sow-badge.error { color: #ff6b6b !important; animation: none !important; }
    @keyframes d-dart-glow-pulse { 0%, 100% { text-shadow: 0 0 5px rgba(255, 153, 0, 0.5) !important; } 50% { text-shadow: 0 0 15px rgba(255, 153, 0, 0.8) !important; } }
    .d-dart-sow-badge:hover { color: #FFD700 !important; animation: none !important; }
    .d-dart-sow-details { max-height: 0 !important; overflow: hidden !important; background: #1a242f !important; border-radius: 4px !important; transition: all 0.3s ease !important; margin-bottom: 0 !important; }
    .d-dart-sow-details.expanded { max-height: 150px !important; margin-bottom: 8px !important; padding: 8px 10px !important; border: 1px solid #FF9900 !important; }
    .d-dart-sow-flex { display: flex !important; flex-wrap: wrap !important; gap: 8px 16px !important; align-items: center !important; }
    .d-dart-sow-item { display: flex !important; align-items: center !important; gap: 4px !important; font-size: 10px !important; white-space: nowrap !important; }
    .d-dart-sow-item-label { color: #888 !important; }
    .d-dart-sow-item-value { color: #FF9900 !important; font-weight: 600 !important; }
    .d-dart-sow-item-value.warning { color: #ffd700 !important; }
    .d-dart-sow-item-value.success { color: #00ff88 !important; }
    .d-dart-sow-item-value.disabled { color: #ff6b6b !important; }
    .d-dart-id-row { display: flex !important; margin-bottom: 8px !important; border-bottom: 1px solid #485769 !important; padding-bottom: 8px !important; }
    .d-dart-id-item { flex: 1 !important; display: flex !important; align-items: center !important; justify-content: center !important; padding: 6px 4px !important; color: #FEBD69 !important; text-decoration: none !important; border-right: 1px solid #485769 !important; transition: all 0.2s ease !important; cursor: pointer !important; }
    .d-dart-id-item:last-child { border-right: none !important; }
    .d-dart-id-item:hover { background: rgba(255,153,0,0.1) !important; }
    .d-dart-id-item:hover .d-dart-id-value { color: #FF9900 !important; text-decoration: underline !important; }
    .d-dart-id-content { display: flex !important; flex-direction: column !important; align-items: center !important; gap: 2px !important; min-width: 0 !important; }
    .d-dart-id-label { font-size: 9px !important; font-weight: 600 !important; color: #888 !important; text-transform: uppercase !important; }
    .d-dart-id-value { display: flex !important; align-items: center !important; gap: 4px !important; font-size: 14px !important; font-weight: 700 !important; color: #FEBD69 !important; overflow: hidden !important; text-overflow: ellipsis !important; white-space: nowrap !important; }
    .d-dart-id-icon { font-size: 14px !important; flex-shrink: 0 !important; }
    .d-dart-lane-row { display: flex !important; align-items: center !important; justify-content: space-between !important; padding: 8px 10px !important; background: #1a242f !important; border-radius: 6px !important; margin-bottom: 8px !important; gap: 8px !important; }
    .d-dart-lane-origin, .d-dart-lane-dest { font-size: 11px !important; font-weight: 600 !important; color: #FFF !important; flex-shrink: 1 !important; min-width: 0 !important; overflow: hidden !important; text-overflow: ellipsis !important; white-space: nowrap !important; max-width: 38% !important; }
    .d-dart-lane-arrow-container { flex: 1 !important; display: flex !important; justify-content: center !important; align-items: center !important; overflow: hidden !important; position: relative !important; min-width: 80px !important; height: 20px !important; }
    .d-dart-lane-arrow-track { position: relative !important; width: 100% !important; height: 100% !important; display: flex !important; align-items: center !important; }
    .d-dart-arrow-moving { position: absolute !important; color: #FFD700 !important; font-size: 14px !important; font-weight: bold !important; animation: d-dart-arrow-slide 1.5s linear infinite !important; text-shadow: 0 0 8px rgba(255,215,0,0.8) !important; }
    @keyframes d-dart-arrow-slide { 0% { left: 0%; opacity: 0; } 10% { opacity: 1; } 90% { opacity: 1; } 100% { left: 100%; opacity: 0; } }
    .d-dart-holds-row { display: flex !important; align-items: center !important; padding: 8px 10px !important; background: #1a242f !important; border-radius: 6px !important; }
    .d-dart-holds-label { flex: 0 0 20% !important; font-size: 11px !important; font-weight: 700 !important; color: #888 !important; text-transform: uppercase !important; }
    .d-dart-hold-item { flex: 0 0 40% !important; font-size: 12px !important; font-weight: 700 !important; text-align: center !important; }
    .d-dart-hold-item.has-hold { color: #00ff88 !important; }
    .d-dart-hold-item.no-hold { color: #ff6b6b !important; }
    .d-dart-section-title { font-size: 11px !important; font-weight: 600 !important; color: #FF9900 !important; text-transform: uppercase !important; letter-spacing: 1px !important; margin-bottom: 8px !important; padding-bottom: 4px !important; border-bottom: 1px solid #37475A !important; }
    .d-dart-stop-card { background: #37475A !important; border-radius: 8px !important; padding: 10px !important; margin-bottom: 6px !important; border-left: 3px solid #FF9900 !important; }
    .d-dart-stop-card:last-child { margin-bottom: 0 !important; }
    .d-dart-stop-header { display: flex !important; justify-content: space-between !important; align-items: center !important; margin-bottom: 10px !important; padding-bottom: 8px !important; border-bottom: 1px solid #485769 !important; }
    .d-dart-stop-title { font-size: 10px !important; font-weight: 700 !important; color: #FFF !important; }
    .d-dart-stop-badges { display: flex !important; gap: 6px !important; }
    .d-dart-stop-badge { font-size: 10px !important; padding: 3px 8px !important; border-radius: 4px !important; font-weight: 600 !important; text-transform: uppercase !important; }
    .d-dart-stop-badge.pickup { background: rgba(255,153,0,0.25) !important; color: #FF9900 !important; }
    .d-dart-stop-badge.dropoff { background: rgba(0,255,136,0.2) !important; color: #00ff88 !important; }
    .d-dart-stop-badge.load-type { background: #485769 !important; color: #FFF !important; }
    .d-dart-timestamp-boxes { display: flex !important; flex-direction: column !important; gap: 6px !important; }
    .d-dart-ts-box { display: flex !important; width: 100% !important; background: #1a242f !important; border-radius: 6px !important; overflow: hidden !important; border: 1px solid #485769 !important; }
    .d-dart-ts-col { padding: 8px 10px !important; display: flex !important; flex-direction: column !important; gap: 4px !important; border-right: 1px solid #485769 !important; }
    .d-dart-ts-col:last-child { border-right: none !important; }
    .d-dart-ts-col.planned { width: 35% !important; }
    .d-dart-ts-col.actual { width: 35% !important; }
    .d-dart-ts-col.delay { width: 30% !important; text-align: center !important; justify-content: center !important; align-items: center !important; }
    .d-dart-ts-col.delay.early { background: rgba(0,255,136,0.15) !important; }
    .d-dart-ts-col.delay.late { background: rgba(255,107,107,0.15) !important; }
    .d-dart-ts-col-label { font-size: 10px !important; font-weight: 600 !important; color: #888 !important; text-transform: uppercase !important; }
    .d-dart-ts-col-value { font-size: 13px !important; font-weight: 700 !important; color: #FFF !important; }
    .d-dart-ts-col.delay.early .d-dart-ts-col-value { color: #00ff88 !important; }
    .d-dart-ts-col.delay.late .d-dart-ts-col-value { color: #ff6b6b !important; }
    .d-dart-borrowed-indicator { font-size: 9px !important; color: #ffd43b !important; background: rgba(255,212,59,0.15) !important; padding: 2px 6px !important; border-radius: 3px !important; margin-top: 4px !important; font-style: italic !important; }
    .d-dart-progress { background: #232F3E !important; border-radius: 8px !important; padding: 15px !important; margin-bottom: 12px !important; }
    .d-dart-progress-title { font-size: 18px !important; font-weight: 700 !important; color: #FF9900 !important; margin-bottom: 12px !important; text-align: center !important; }
    .d-dart-progress-steps { display: flex !important; flex-direction: column !important; gap: 6px !important; }
    .d-dart-step { display: flex !important; align-items: center !important; gap: 10px !important; padding: 8px 12px !important; background: #37475A !important; border-radius: 6px !important; font-size: 13px !important; transition: all 0.3s ease !important; }
    .d-dart-step.pending { opacity: 0.5 !important; }
    .d-dart-step.active { background: #485769 !important; border-left: 3px solid #FF9900 !important; animation: d-dart-pulse 1s infinite !important; }
    .d-dart-step.completed { background: rgba(0,255,136,0.1) !important; border-left: 3px solid #00ff88 !important; }
    .d-dart-step.error { background: rgba(255,107,107,0.1) !important; border-left: 3px solid #ff6b6b !important; }
    .d-dart-step-icon { font-size: 16px !important; width: 24px !important; text-align: center !important; }
    .d-dart-step-text { flex: 1 !important; font-weight: 500 !important; }
    .d-dart-step-status { font-size: 10px !important; color: #888 !important; }
    .d-dart-enhanced-progress { background: #232F3E !important; border-radius: 8px !important; padding: 15px !important; }
    .d-dart-progress-header { display: flex !important; justify-content: space-between !important; align-items: center !important; margin-bottom: 15px !important; }
    .d-dart-progress-controls { display: flex !important; gap: 6px !important; }
    .d-dart-control-btn { background: #37475A !important; border: 1px solid #485769 !important; border-radius: 6px !important; padding: 6px 10px !important; font-size: 14px !important; cursor: pointer !important; transition: all 0.2s !important; color: #FFF !important; }
    .d-dart-control-btn:hover { background: #485769 !important; border-color: #FF9900 !important; }
    .d-dart-control-btn.cancel:hover { background: rgba(255,107,107,0.2) !important; border-color: #ff6b6b !important; }
    .d-dart-progress-stats { display: grid !important; grid-template-columns: repeat(4, 1fr) !important; gap: 8px !important; margin-bottom: 15px !important; }
    .d-dart-stat { background: #37475A !important; border-radius: 6px !important; padding: 10px 8px !important; text-align: center !important; }
    .d-dart-stat.success .d-dart-stat-value { color: #00ff88 !important; }
    .d-dart-stat.error .d-dart-stat-value { color: #ff6b6b !important; }
    .d-dart-stat-value { display: block !important; font-size: 22px !important; font-weight: 700 !important; color: #FF9900 !important; }
    .d-dart-stat-label { font-size: 9px !important; color: #888 !important; text-transform: uppercase !important; }
    .d-dart-progress-bar-container { height: 8px !important; background: #37475A !important; border-radius: 4px !important; overflow: hidden !important; margin-bottom: 10px !important; }
    .d-dart-progress-bar { height: 100% !important; background: linear-gradient(90deg, #FF9900, #00ff88) !important; border-radius: 4px !important; transition: width 0.3s ease !important; }
    .d-dart-progress-info { text-align: center !important; margin-bottom: 10px !important; }
    .d-dart-progress-status { color: #FFF !important; font-size: 12px !important; margin-bottom: 4px !important; }
    .d-dart-progress-eta { color: #888 !important; font-size: 11px !important; }
    .d-dart-progress-footer { display: flex !important; justify-content: space-between !important; font-size: 10px !important; color: #666 !important; }
    .d-dart-progress-token .token-ok { color: #00ff88 !important; }
    .d-dart-progress-token .token-warning { color: #ffd700 !important; }
    .d-dart-progress-token .token-error { color: #ff6b6b !important; }
    .d-dart-batch-summary { display: grid !important; grid-template-columns: repeat(9, 1fr) !important; gap: 6px !important; margin-bottom: 10px !important; }
    .d-dart-batch-stat { background: #37475A !important; border-radius: 6px !important; padding: 8px 4px !important; text-align: center !important; font-size: 10px !important; font-weight: 600 !important; }
    .d-dart-batch-stat.success { color: #00ff88 !important; }
    .d-dart-batch-stat.recovered { color: #FF9900 !important; }
    .d-dart-batch-stat.released { color: #4dabf7 !important; }
    .d-dart-batch-stat.analysis { color: #4dabf7 !important; }
    .d-dart-batch-stat.pending { color: #ffd700 !important; }
    .d-dart-batch-stat.error { color: #ff6b6b !important; }
    .d-dart-batch-stat.hold-created { color: #4dabf7 !important; }
    .d-dart-batch-table-container { max-height: 300px !important; overflow-y: auto !important; margin-bottom: 10px !important; border-radius: 6px !important; border: 1px solid #37475A !important; }
    .d-dart-batch-table { width: 100% !important; border-collapse: collapse !important; font-size: 10px !important; }
    .d-dart-batch-table th { background: #37475A !important; color: #888 !important; padding: 6px 4px !important; text-align: left !important; font-weight: 600 !important; text-transform: uppercase !important; position: sticky !important; top: 0 !important; font-size: 9px !important; }
    .d-dart-batch-table td { padding: 5px 4px !important; border-bottom: 1px solid #37475A !important; color: #FFF !important; }
    .d-dart-batch-table tr.success td { background: rgba(0,255,136,0.05) !important; }
    .d-dart-batch-table tr.pending td { background: rgba(255,215,0,0.05) !important; }
    .d-dart-batch-table tr.error td { background: rgba(255,107,107,0.05) !important; }
    .d-dart-batch-table tr.recovered td { background: rgba(255,153,0,0.05) !important; }
    .d-dart-batch-table tr.analysis td { background: rgba(77,171,247,0.05) !important; }
    .d-dart-batch-table tr.hold-created td { background: rgba(77,171,247,0.05) !important; }
    .d-dart-order-cell { color: #FEBD69 !important; }
    .d-dart-download-buttons { display: flex !important; gap: 8px !important; justify-content: center !important; }
    .d-dart-download-btn { padding: 8px 16px !important; background: #37475A !important; border: 1px solid #485769 !important; border-radius: 6px !important; color: #FFF !important; font-size: 11px !important; font-weight: 600 !important; cursor: pointer !important; transition: all 0.2s ease !important; }
    .d-dart-download-btn:hover { background: #485769 !important; border-color: #FF9900 !important; }
    .d-dart-error { background: rgba(255,107,107,0.15) !important; border: 1px solid #ff6b6b !important; border-radius: 8px !important; padding: 20px !important; color: #ff6b6b !important; text-align: center !important; }
    .d-dart-error-icon { font-size: 32px !important; margin-bottom: 8px !important; }
    .d-dart-error-message { font-size: 12px !important; line-height: 1.4 !important; }
    .d-dart-empty { text-align: center !important; padding: 20px !important; color: #888 !important; font-size: 13px !important; }
    .d-dart-sow-error { background: rgba(255,107,107,0.1) !important; border: 1px solid #ff6b6b !important; border-radius: 8px !important; padding: 20px !important; text-align: center !important; margin-bottom: 12px !important; }
    .d-dart-sow-error-icon { font-size: 40px !important; margin-bottom: 10px !important; }
    .d-dart-sow-error-title { font-size: 16px !important; font-weight: 700 !important; color: #ff6b6b !important; margin-bottom: 8px !important; }
    .d-dart-sow-error-message { font-size: 12px !important; color: #ccc !important; margin-bottom: 15px !important; }
    .d-dart-sow-error-instructions { font-size: 11px !important; color: #888 !important; text-align: left !important; margin-bottom: 15px !important; padding: 10px !important; background: #232F3E !important; border-radius: 6px !important; }
    .d-dart-sow-error-instructions p { margin: 5px 0 !important; }
    .d-dart-sow-login-btn { display: inline-block !important; padding: 10px 20px !important; background: #37475A !important; color: #FFF !important; text-decoration: none !important; border-radius: 6px !important; font-weight: 600 !important; margin-bottom: 10px !important; transition: all 0.2s ease !important; }
    .d-dart-sow-login-btn:hover { background: #485769 !important; }
    .d-dart-sow-retry-btn { display: block !important; width: 100% !important; padding: 10px !important; background: #FF9900 !important; border: none !important; border-radius: 6px !important; color: #232F3E !important; font-weight: 700 !important; cursor: pointer !important; transition: all 0.2s ease !important; }
    .d-dart-sow-retry-btn:hover { background: #FEBD69 !important; }
    #d-dart-approval-overlay { position: fixed !important; top: 0 !important; left: 0 !important; width: 100% !important; height: 100% !important; background: rgba(0,0,0,0.7) !important; display: flex !important; align-items: center !important; justify-content: center !important; z-index: 2147483650 !important; animation: d-dart-fadeIn 0.2s ease !important; }
    .d-dart-approval-popup { background: #232F3E !important; border: 2px solid #FF9900 !important; border-radius: 12px !important; width: 420px !important; max-width: 90vw !important; box-shadow: 0 20px 60px rgba(0,0,0,0.5) !important; animation: d-dart-popupSlide 0.3s ease !important; }
    .d-dart-approval-header { background: linear-gradient(90deg, #FF9900, #E88B00) !important; padding: 12px 16px !important; border-radius: 10px 10px 0 0 !important; display: flex !important; justify-content: space-between !important; align-items: center !important; }
    .d-dart-approval-title { font-size: 16px !important; font-weight: 700 !important; color: #232F3E !important; }
    .d-dart-approval-timer { font-size: 14px !important; font-weight: 700 !important; color: #232F3E !important; background: rgba(255,255,255,0.3) !important; padding: 4px 10px !important; border-radius: 20px !important; transition: all 0.3s ease !important; }
    .d-dart-approval-timer.warning { background: rgba(255,200,0,0.5) !important; color: #000 !important; }
    .d-dart-approval-timer.critical { background: rgba(255,100,100,0.8) !important; color: #FFF !important; animation: d-dart-pulse 0.5s infinite !important; }
    .d-dart-approval-body { padding: 20px !important; }
    .d-dart-approval-order-id, .d-dart-approval-shipper, .d-dart-approval-charge-confirm { display: flex !important; justify-content: space-between !important; align-items: center !important; padding: 8px 12px !important; background: #37475A !important; border-radius: 6px !important; margin-bottom: 10px !important; }
    .d-dart-approval-label { font-size: 12px !important; color: #888 !important; }
    .d-dart-approval-value { font-size: 14px !important; font-weight: 600 !important; color: #FEBD69 !important; }
    .d-dart-approval-charge-info { background: #1a242f !important; border: 1px solid #485769 !important; border-radius: 8px !important; padding: 15px !important; margin-bottom: 15px !important; }
    .d-dart-approval-charge-title { font-size: 12px !important; color: #888 !important; margin-bottom: 10px !important; text-align: center !important; }
    .d-dart-approval-charge-details { margin-bottom: 10px !important; }
    .d-dart-approval-stop-line { font-size: 13px !important; color: #FFF !important; padding: 6px 0 !important; border-bottom: 1px solid #37475A !important; }
    .d-dart-approval-stop-line:last-child { border-bottom: none !important; }
    .d-dart-approval-stop-line strong { color: #ffd700 !important; }
    .d-dart-approval-total { display: flex !important; justify-content: space-between !important; align-items: center !important; padding-top: 10px !important; border-top: 2px solid #FF9900 !important; margin-top: 10px !important; }
    .d-dart-approval-total-label { font-size: 14px !important; font-weight: 600 !important; color: #FFF !important; }
    .d-dart-approval-total-value { font-size: 20px !important; font-weight: 700 !important; color: #ffd700 !important; }
    .d-dart-approval-question { font-size: 14px !important; color: #FFF !important; text-align: center !important; margin-bottom: 5px !important; }
    .d-dart-approval-buttons { display: flex !important; gap: 10px !important; padding: 15px 20px 20px !important; justify-content: center !important; }
    .d-dart-approval-btn { padding: 12px 24px !important; border: none !important; border-radius: 8px !important; font-size: 14px !important; font-weight: 700 !important; cursor: pointer !important; transition: all 0.2s ease !important; min-width: 100px !important; }
    .d-dart-approval-btn:hover { transform: translateY(-2px) !important; box-shadow: 0 4px 15px rgba(0,0,0,0.3) !important; }
    .d-dart-approval-btn.yes { background: #00ff88 !important; color: #232F3E !important; }
    .d-dart-approval-btn.yes:hover { background: #00cc6a !important; }
    .d-dart-approval-btn.no { background: #ff6b6b !important; color: #FFF !important; }
    .d-dart-approval-btn.no:hover { background: #ff4444 !important; }
    .d-dart-approval-btn.skip { background: #37475A !important; color: #FFF !important; border: 1px solid #485769 !important; }
    .d-dart-approval-btn.skip:hover { background: #485769 !important; border-color: #FF9900 !important; }
    .d-dart-approval-btn.submit { background: #FF9900 !important; color: #232F3E !important; }
    .d-dart-approval-btn.submit:hover { background: #FEBD69 !important; }
    .d-dart-approval-btn.cancel { background: #37475A !important; color: #FFF !important; }
    .d-dart-approval-btn.cancel:hover { background: #485769 !important; }
    .d-dart-approval-auth-section { margin-top: 10px !important; }
    .d-dart-approval-auth-label { display: block !important; font-size: 12px !important; color: #888 !important; margin-bottom: 8px !important; }
    .d-dart-approval-auth-input { width: 100% !important; padding: 12px !important; border: 2px solid #37475A !important; border-radius: 6px !important; background: #1a242f !important; color: #FFF !important; font-size: 14px !important; outline: none !important; transition: border-color 0.2s ease !important; }
    .d-dart-approval-auth-input:focus { border-color: #FF9900 !important; }
    .d-dart-approval-auth-input.error { border-color: #ff6b6b !important; }
    .d-dart-approval-auth-input::placeholder { color: #666 !important; }
    .d-dart-approval-auth-error { color: #ff6b6b !important; font-size: 11px !important; margin-top: 5px !important; min-height: 16px !important; }
    #d-dart-resume-overlay { position: fixed !important; top: 0 !important; left: 0 !important; width: 100% !important; height: 100% !important; background: rgba(0,0,0,0.7) !important; display: flex !important; align-items: center !important; justify-content: center !important; z-index: 2147483650 !important; }
    .d-dart-resume-popup { background: #232F3E !important; border: 2px solid #FF9900 !important; border-radius: 12px !important; padding: 25px !important; text-align: center !important; max-width: 400px !important; }
    .d-dart-resume-title { font-size: 18px !important; font-weight: 700 !important; color: #FF9900 !important; margin-bottom: 15px !important; }
    .d-dart-resume-info { color: #FFF !important; margin-bottom: 20px !important; line-height: 1.6 !important; }
    .d-dart-resume-buttons { display: flex !important; gap: 10px !important; justify-content: center !important; }
    .d-dart-resume-btn { padding: 10px 20px !important; border: none !important; border-radius: 6px !important; font-weight: 600 !important; cursor: pointer !important; transition: all 0.2s ease !important; }
    .d-dart-resume-btn:hover { transform: translateY(-2px) !important; }
    .d-dart-resume-btn.yes { background: #00ff88 !important; color: #232F3E !important; }
    .d-dart-resume-btn.no { background: #37475A !important; color: #FFF !important; }
    .d-dart-copy-popup { position: fixed !important; top: 50% !important; left: 50% !important; transform: translate(-50%, -50%) !important; background: #00ff88 !important; color: #232F3E !important; padding: 15px 30px !important; border-radius: 10px !important; font-size: 16px !important; font-weight: 700 !important; z-index: 2147483649 !important; box-shadow: 0 8px 30px rgba(0,0,0,0.4) !important; animation: d-dart-pop 0.3s ease !important; }
    #d-dart-body::-webkit-scrollbar { width: 5px !important; }
    #d-dart-body::-webkit-scrollbar-track { background: #1a242f !important; }
    #d-dart-body::-webkit-scrollbar-thumb { background: #37475A !important; border-radius: 3px !important; }
    #d-dart-body::-webkit-scrollbar-thumb:hover { background: #FF9900 !important; }
    .d-dart-batch-table-container::-webkit-scrollbar { width: 4px !important; }
    .d-dart-batch-table-container::-webkit-scrollbar-track { background: #1a242f !important; }
    .d-dart-batch-table-container::-webkit-scrollbar-thumb { background: #37475A !important; border-radius: 2px !important; }
    .d-dart-shippers-list::-webkit-scrollbar { width: 4px !important; }
    .d-dart-shippers-list::-webkit-scrollbar-track { background: #1a242f !important; }
    .d-dart-shippers-list::-webkit-scrollbar-thumb { background: #485769 !important; border-radius: 2px !important; }
`;

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 32: UI CONTROLLER (Fix #3, #7, #9, #11)
     * ═══════════════════════════════════════════════════════════════════════════ */

    const UIController = (() => {
        let dom = null;
        let toastTimeout = null;
        let dragState = { isDragging: false, offsetX: 0, offsetY: 0 };
        let boundHandlers = new Map();
        let debouncedAnalyze = null;
        // Fix #3: stored escape handler
        let settingsEscapeHandler = null;
        // Fix #7: stored keyboard shortcut handler
        let keyboardShortcutHandler = null;
        // Fix #11: cached progress DOM refs
        let progressDOMRefs = null;

        const injectStyles = () => { GM_addStyle(Styles); };

        const createDOM = () => {
            const container = document.createElement('div');
            container.id = 'd-dart';
            container.setAttribute('role', 'application');
            container.setAttribute('aria-label', Messages.ACCESSIBILITY.PANEL_LABEL);
            container.innerHTML = `
                <div id="d-dart-minimized-icon" title="Expand D-DART" role="button" tabindex="0" aria-label="${Messages.ACCESSIBILITY.EXPAND_PANEL}">🚛</div>
                <div id="d-dart-header">
                    <div class="d-dart-header-row-1">
                        <h3 class="d-dart-app-title">🚛 D-DART</h3>
                        <div class="d-dart-header-buttons">
                            <button class="d-dart-header-btn" id="d-dart-settings-btn" title="Settings" aria-label="${Messages.ACCESSIBILITY.OPEN_SETTINGS}">⚙️</button>
                            <button class="d-dart-header-btn" id="d-dart-reset-btn" title="Reset form" aria-label="${Messages.ACCESSIBILITY.RESET_FORM}">↻</button>
                            <button class="d-dart-header-btn" id="d-dart-debug-btn" title="Copy Debug Log" aria-label="${Messages.ACCESSIBILITY.COPY_DEBUG}">🔍</button>
                            <button class="d-dart-header-btn" id="d-dart-toggle" title="Minimize" aria-label="${Messages.ACCESSIBILITY.MINIMIZE_PANEL}">−</button>
                        </div>
                    </div>
                    <div class="d-dart-header-row-2">
                        <div class="d-dart-mode-toggle" id="d-dart-mode-toggle" role="switch" tabindex="0" aria-label="${Messages.ACCESSIBILITY.TOGGLE_MODE}">
                            <span class="d-dart-mode-option action" id="d-dart-mode-action">⚡ACTION</span>
                            <span class="d-dart-mode-option review" id="d-dart-mode-review">📋REVIEW</span>
                        </div>
                        <div class="d-dart-signature-block">
                            <span class="d-dart-made-by">Made by</span>
                            <span class="d-dart-signature">${CONFIG.AUTHOR}</span>
                        </div>
                    </div>
                </div>
                <div id="d-dart-body">
                    <div class="d-dart-input-group">
                        <input type="text" class="d-dart-input" id="d-dart-order-id" placeholder="Enter Order ID(s) - supports ${CONFIG.BATCH.MAX_ORDERS_PER_SESSION}+ orders" autocomplete="off" spellcheck="false" aria-label="${Messages.ACCESSIBILITY.ORDER_INPUT}" aria-describedby="d-dart-input-hint">
                        <span id="d-dart-input-hint" class="d-dart-visually-hidden">Enter one or more order IDs separated by commas, spaces, or new lines</span>
                        <button class="d-dart-btn" id="d-dart-analyze-btn" aria-label="${Messages.ACCESSIBILITY.ANALYZE_ORDERS}">Analyze</button>
                    </div>
                    <div id="d-dart-results-container" role="region" aria-label="${Messages.ACCESSIBILITY.RESULTS_REGION}" aria-live="polite"></div>
                </div>
            `;
            document.body.appendChild(container);
            const toast = document.createElement('div');
            toast.className = 'd-dart-toast'; toast.id = 'd-dart-toast';
            toast.setAttribute('role', 'alert'); toast.setAttribute('aria-live', 'assertive');
            document.body.appendChild(toast);
            dom = {
                container, body: document.getElementById('d-dart-body'),
                header: document.getElementById('d-dart-header'),
                toggle: document.getElementById('d-dart-toggle'),
                reset: document.getElementById('d-dart-reset-btn'),
                settings: document.getElementById('d-dart-settings-btn'),
                input: document.getElementById('d-dart-order-id'),
                analyzeBtn: document.getElementById('d-dart-analyze-btn'),
                results: document.getElementById('d-dart-results-container'),
                minimizedIcon: document.getElementById('d-dart-minimized-icon'),
                debugBtn: document.getElementById('d-dart-debug-btn'),
                modeToggle: document.getElementById('d-dart-mode-toggle'),
                modeAction: document.getElementById('d-dart-mode-action'),
                modeReview: document.getElementById('d-dart-mode-review'),
                toast
            };
        };

        const setupEventListeners = () => {
            dom.toggle.addEventListener('click', () => setMinimized(true));
            dom.minimizedIcon.addEventListener('click', (e) => { e.stopPropagation(); setMinimized(false); });
            dom.minimizedIcon.addEventListener('keypress', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setMinimized(false); } });
            dom.container.addEventListener('click', (e) => { if (AppState.get('isMinimized') && e.target === dom.container) setMinimized(false); });
            dom.reset.addEventListener('click', resetForm);
            dom.settings.addEventListener('click', openSettings);
            dom.input.addEventListener('keypress', (e) => { if (e.key === 'Enter' && !AppState.get('isProcessing')) debouncedAnalyze(); });
            dom.input.addEventListener('input', () => { dom.input.classList.remove('error'); dom.input.removeAttribute('aria-invalid'); });
            dom.analyzeBtn.addEventListener('click', () => { if (!AppState.get('isProcessing')) debouncedAnalyze(); });
            dom.debugBtn.addEventListener('click', copyDebugLog);
            dom.results.addEventListener('click', handleResultsClick);
            dom.modeToggle.addEventListener('click', () => { ModeManager.toggle(); });
            dom.modeToggle.addEventListener('keypress', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); ModeManager.toggle(); } });
            setupDragging();
            if (CONFIG.FEATURES.KEYBOARD_SHORTCUTS) setupKeyboardShortcuts();
        };

        const handleResultsClick = (e) => {
            const toggleEl = e.target.closest('[data-toggle-target]');
            if (toggleEl) {
                const target = document.getElementById(toggleEl.dataset.toggleTarget);
                if (target) { const isExpanded = target.classList.toggle('expanded'); toggleEl.setAttribute('aria-expanded', isExpanded); target.setAttribute('aria-hidden', !isExpanded); }
                return;
            }
            if (e.target.id === 'd-dart-download-csv') { downloadCSV(); return; }
            if (e.target.id === 'd-dart-download-txt') { downloadTXT(); return; }
            if (e.target.id === 'd-dart-sow-retry') { SOWConfigManager.fetch(); return; }
            if (e.target.id === 'd-dart-pause-btn') {
                BatchProcessor.pause();
                const pauseBtn = document.getElementById('d-dart-pause-btn');
                const resumeBtn = document.getElementById('d-dart-resume-btn');
                if (pauseBtn) pauseBtn.style.display = 'none';
                if (resumeBtn) resumeBtn.style.display = 'inline-block';
                return;
            }
            if (e.target.id === 'd-dart-resume-btn') {
                BatchProcessor.resume();
                const pauseBtn = document.getElementById('d-dart-pause-btn');
                const resumeBtn = document.getElementById('d-dart-resume-btn');
                if (resumeBtn) resumeBtn.style.display = 'none';
                if (pauseBtn) pauseBtn.style.display = 'inline-block';
                return;
            }
            if (e.target.id === 'd-dart-cancel-btn') {
                if (confirm('Are you sure you want to cancel? Progress will be saved.')) BatchProcessor.cancel();
                return;
            }
        };

        // Fix #7: store handler for cleanup in destroy()
        const setupKeyboardShortcuts = () => {
            keyboardShortcutHandler = (e) => {
                if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
                if ((e.ctrlKey || e.metaKey) && e.shiftKey) {
                    switch (e.code) {
                        case CONFIG.KEYBOARD.TOGGLE_MINIMIZE: e.preventDefault(); setMinimized(!AppState.get('isMinimized')); break;
                        case CONFIG.KEYBOARD.SETTINGS: e.preventDefault(); if (AppState.get('isSettingsOpen')) closeSettings(); else openSettings(); break;
                        case CONFIG.KEYBOARD.DEBUG: e.preventDefault(); copyDebugLog(); break;
                    }
                }
            };
            document.addEventListener('keydown', keyboardShortcutHandler);
        };

        const setupDragging = () => {
            const startDrag = (e) => {
                const ignoredSelectors = ['.d-dart-header-buttons', '.d-dart-input', '.d-dart-btn', '.d-dart-download-btn', '.d-dart-debug-btn', '.d-dart-banner-toggle', '.d-dart-sow-badge', '.d-dart-id-item', '[data-toggle-target]', 'a', '.d-dart-control-btn', '.d-dart-sow-retry-btn', '.d-dart-sow-login-btn', '.d-dart-back-btn', '.d-dart-mode-toggle', '.d-dart-refresh-sow-btn'];
                for (const selector of ignoredSelectors) { if (e.target.closest(selector)) return; }
                dragState.isDragging = true; dom.container.classList.add('dragging');
                const clientX = e.clientX || e.touches?.[0]?.clientX;
                const clientY = e.clientY || e.touches?.[0]?.clientY;
                dragState.offsetX = clientX - dom.container.offsetLeft;
                dragState.offsetY = clientY - dom.container.offsetTop;
                e.preventDefault();
            };
            const moveDrag = (e) => {
                if (!dragState.isDragging) return;
                e.preventDefault();
                const clientX = e.clientX || e.touches?.[0]?.clientX;
                const clientY = e.clientY || e.touches?.[0]?.clientY;
                requestAnimationFrame(() => {
                    let newX = clientX - dragState.offsetX; let newY = clientY - dragState.offsetY;
                    const maxX = window.innerWidth - dom.container.offsetWidth;
                    const maxY = window.innerHeight - dom.container.offsetHeight;
                    newX = Math.max(0, Math.min(newX, maxX)); newY = Math.max(0, Math.min(newY, maxY));
                    dom.container.style.setProperty('left', `${newX}px`, 'important');
                    dom.container.style.setProperty('top', `${newY}px`, 'important');
                    dom.container.style.setProperty('right', 'auto', 'important');
                });
            };
            const endDrag = () => { dragState.isDragging = false; dom.container.classList.remove('dragging'); };
            boundHandlers.set('moveDrag', moveDrag); boundHandlers.set('endDrag', endDrag);
            dom.header.addEventListener('mousedown', startDrag);
            dom.container.addEventListener('mousedown', (e) => { if (AppState.get('isMinimized')) startDrag(e); });
            document.addEventListener('mousemove', moveDrag); document.addEventListener('mouseup', endDrag);
            dom.header.addEventListener('touchstart', startDrag, { passive: false });
            dom.container.addEventListener('touchstart', (e) => { if (AppState.get('isMinimized')) startDrag(e); }, { passive: false });
            document.addEventListener('touchmove', moveDrag, { passive: false }); document.addEventListener('touchend', endDrag);
        };

        const setupStateSubscriptions = () => {
            AppState.subscribe('isProcessing', (isProcessing) => {
                dom.analyzeBtn.disabled = isProcessing || !SOWConfigManager.isLoaded();
                dom.analyzeBtn.textContent = isProcessing ? 'Processing...' : 'Analyze';
                dom.analyzeBtn.setAttribute('aria-busy', isProcessing);
                dom.analyzeBtn.classList.toggle('loading', isProcessing);
                dom.input.disabled = isProcessing;
            });
            AppState.subscribe('isMinimized', updateMinimizedState);
            AppState.subscribe('sowStatus', () => { updateSOWIndicator(); updateAnalyzeButtonState(); });
            AppState.subscribe('operationMode', updateModeIndicator);
        };

        const updateAnalyzeButtonState = () => {
            const isProcessing = AppState.get('isProcessing');
            const sowLoaded = SOWConfigManager.isLoaded();
            dom.analyzeBtn.disabled = isProcessing || !sowLoaded;
            if (!sowLoaded && !isProcessing) dom.analyzeBtn.title = 'SOW not loaded - click refresh';
            else dom.analyzeBtn.title = '';
        };

        const setMinimized = (minimized) => {
            AppState.set('isMinimized', minimized);
            dom.container.classList.toggle('minimized', minimized);
            if (!minimized) setTimeout(() => dom.input.focus(), 100);
            Telemetry.track(TelemetryEventType.USER_ACTION, { action: minimized ? 'minimize' : 'expand' });
        };

        const updateMinimizedState = () => {
            const isHealthy = HealthCheck.isHealthy();
            dom.container.classList.remove('healthy', 'unhealthy');
            dom.container.classList.add(isHealthy ? 'healthy' : 'unhealthy');
        };

        const updateModeIndicator = () => {
            const mode = ModeManager.getMode();
            if (dom?.modeAction && dom?.modeReview) {
                dom.modeAction.classList.toggle('active', mode === CONFIG.MODE.ACTION);
                dom.modeReview.classList.toggle('active', mode === CONFIG.MODE.REVIEW);
                dom.modeToggle.setAttribute('aria-checked', mode === CONFIG.MODE.ACTION ? 'true' : 'false');
            }
        };

        const startAnalysis = () => {
            if (AppState.get('isSettingsOpen')) closeSettings();
            if (!SOWConfigManager.isLoaded()) { showToast(Messages.ERRORS.SOW_SERVER_UNREACHABLE, 'error'); return; }
            const input = dom.input.value.trim();
            const validation = Validator.parseOrderIds(input);
            dom.input.classList.remove('error'); dom.input.removeAttribute('aria-invalid');
            if (!validation.valid || validation.sanitized.length === 0) {
                dom.input.classList.add('error'); dom.input.setAttribute('aria-invalid', 'true'); dom.input.focus();
                const errorMsg = validation.errors.length > 0 ? validation.errors[0] : Messages.ERRORS.INVALID_ORDER_IDS;
                showToast(errorMsg, 'error'); return;
            }
            if (validation.errors.length > 0) { Logger.warn('Some order IDs were invalid', validation.errors); showToast(`${validation.sanitized.length} valid IDs found, ${validation.errors.length} invalid`, 'warning'); }
            if (validation.duplicatesRemoved > 0) Logger.info(`Removed ${validation.duplicatesRemoved} duplicate order IDs`);
            AppState.set('isProcessing', true);
            BatchProcessor.processBatch(validation.sanitized).finally(() => { AppState.set('isProcessing', false); });
        };

        const resetForm = () => {
            if (AppState.get('isSettingsOpen')) closeSettings();
            dom.input.value = ''; dom.input.classList.remove('error'); dom.input.removeAttribute('aria-invalid');
            dom.results.innerHTML = ''; AppState.resetBatch(); progressDOMRefs = null; dom.input.focus();
            Telemetry.track(TelemetryEventType.USER_ACTION, { action: 'reset' });
        };

        const openSettings = () => {
            if (AppState.get('isSettingsOpen')) {
                const settingsEl = document.getElementById('d-dart-settings-inline');
                if (settingsEl) return;
                Logger.warn('Settings state was stale (DOM missing), re-opening');
                AppState.set('isSettingsOpen', false);
            }
            AppState.set('isSettingsOpen', true);
            const currentContent = dom.results.innerHTML;
            AppState.set('_previousResultsContent', currentContent);
            dom.results.innerHTML = HTMLGenerator.settingsPanelInline();
            const inputGroup = dom.body.querySelector('.d-dart-input-group');
            if (inputGroup) inputGroup.style.display = 'none';
            setupSettingsEventListeners();
            Logger.info('Settings panel opened (inline)');
            Telemetry.track(TelemetryEventType.USER_ACTION, { action: 'open_settings' });
        };

        // Fix #3: remove escape handler on close
        const closeSettings = () => {
            if (!AppState.get('isSettingsOpen')) return;
            AppState.set('isSettingsOpen', false);
            // Clean up escape handler
            if (settingsEscapeHandler) {
                document.removeEventListener('keydown', settingsEscapeHandler);
                settingsEscapeHandler = null;
            }
            const inputGroup = dom.body.querySelector('.d-dart-input-group');
            if (inputGroup) inputGroup.style.display = 'flex';
            const previousContent = AppState.get('_previousResultsContent') || '';
            dom.results.innerHTML = previousContent;
            AppState.set('_previousResultsContent', '');
            Logger.info('Settings panel closed');
        };

        // Fix #3: remove old escape handler before adding new one
        const setupSettingsEventListeners = () => {
            const backBtn = document.getElementById('d-dart-settings-back');
            if (backBtn) backBtn.addEventListener('click', closeSettings);

            const refreshBtn = document.getElementById('d-dart-refresh-sow');
            if (refreshBtn) {
                refreshBtn.addEventListener('click', async () => {
                    refreshBtn.classList.add('loading');
                    refreshBtn.disabled = true;
                    const elapsedEl = document.getElementById('d-dart-refresh-elapsed');
                    const shippersList = document.getElementById('d-dart-shippers-list');
                    const shippersCount = document.getElementById('d-dart-shippers-count');
                    if (shippersList) shippersList.innerHTML = HTMLGenerator.shippersLoadingSkeleton();
                    const startTime = Date.now();
                    let elapsedInterval = null;
                    if (elapsedEl) {
                        elapsedEl.textContent = '⏳ Refreshing... 0s';
                        elapsedInterval = setInterval(() => {
                            const elapsed = Math.round((Date.now() - startTime) / 1000);
                            if (elapsedEl) elapsedEl.textContent = `⏳ Refreshing... ${elapsed}s`;
                        }, 1000);
                    }
                    showToast('🔄 Refreshing SOW configuration...', 'info');
                    try {
                        await SOWConfigManager.fetch();
                        const elapsed = Math.round((Date.now() - startTime) / 1000);
                        if (elapsedEl) elapsedEl.textContent = `✅ Refreshed in ${elapsed}s — Last: Just now`;
                        if (shippersList) shippersList.innerHTML = HTMLGenerator.renderShippersList();
                        const stats = SOWConfigManager.getStatistics();
                        if (shippersCount) shippersCount.textContent = `📋 SHIPPERS (${stats.total})`;
                    } catch (error) {
                        if (elapsedEl) elapsedEl.textContent = `❌ Refresh failed — ${error.message || 'Unknown error'}`;
                        if (shippersList) shippersList.innerHTML = HTMLGenerator.renderShippersList();
                    } finally {
                        if (elapsedInterval) clearInterval(elapsedInterval);
                        refreshBtn.classList.remove('loading');
                        refreshBtn.disabled = false;
                    }
                });
            }

            // Fix #9: toggle single card instead of full re-render
            const shippersList = document.getElementById('d-dart-shippers-list');
            if (shippersList) {
                shippersList.addEventListener('click', (e) => {
                    const header = e.target.closest('[data-toggle-shipper]');
                    if (header) {
                        const shipperName = header.dataset.toggleShipper;
                        const safeId = shipperName.replace(/[^a-zA-Z0-9-]/g, '-');
                        const details = document.getElementById(`d-dart-details-${safeId}`);
                        const btn = header.querySelector('.d-dart-expand-btn');
                        if (details) {
                            const isExpanded = details.classList.toggle('expanded');
                            if (btn) btn.textContent = isExpanded ? '▲' : '▼';
                            if (btn) btn.setAttribute('aria-label', isExpanded ? 'Collapse details' : 'Expand details');
                        }
                        AppState.toggleShipperExpanded(shipperName);
                    }
                });
            }

            // Fix #3: remove existing escape handler before adding
            if (settingsEscapeHandler) {
                document.removeEventListener('keydown', settingsEscapeHandler);
            }
            settingsEscapeHandler = (e) => {
                if (e.key === 'Escape' && AppState.get('isSettingsOpen')) { closeSettings(); }
            };
            document.addEventListener('keydown', settingsEscapeHandler);
        };

        const copyDebugLog = async () => {
            const report = Logger.generateReport();
            try { await Helpers.copyToClipboard(report); showCopyPopup(); }
            catch (e) { showToast(Messages.ERRORS.COPY_FAILED, 'error'); }
        };

        const showCopyPopup = () => {
            const popup = document.createElement('div');
            popup.className = 'd-dart-copy-popup'; popup.textContent = `✅ ${Messages.SUCCESS.DEBUG_COPIED}`;
            popup.setAttribute('role', 'alert'); document.body.appendChild(popup);
            setTimeout(() => popup.remove(), CONFIG.UI.COPY_POPUP_DURATION);
        };

        const downloadCSV = () => {
            const data = AppState.get('batchReportData');
            if (!data || data.length === 0) { showToast(Messages.ERRORS.NO_DATA, 'error'); return; }
            const csv = ReportGenerator.generateCSV(data);
            Helpers.downloadFile(csv, `D-DART_Report_${new Date().toISOString().slice(0, 10)}.csv`, 'text/csv');
            showToast(Messages.SUCCESS.CSV_DOWNLOADED, 'success');
        };

        const downloadTXT = () => {
            const data = AppState.get('batchReportData');
            if (!data || data.length === 0) { showToast(Messages.ERRORS.NO_DATA, 'error'); return; }
            const txt = ReportGenerator.generateTXT(data);
            Helpers.downloadFile(txt, `D-DART_Report_${new Date().toISOString().slice(0, 10)}.txt`, 'text/plain');
            showToast(Messages.SUCCESS.TXT_DOWNLOADED, 'success');
        };

        const checkSOWErrorDisplay = () => {
            if (AppState.get('isSettingsOpen')) return;
            const sowStatus = AppState.get('sowStatus');
            const sowError = AppState.get('sowLastError');
            if (sowStatus === SOWStatus.ERROR || sowStatus === SOWStatus.AUTH_REQUIRED) {
                const isAuthError = sowStatus === SOWStatus.AUTH_REQUIRED;
                dom.results.innerHTML = HTMLGenerator.sowErrorDisplay(sowError || Messages.ERRORS.SOW_SERVER_UNREACHABLE, isAuthError);
            }
        };

        const showToast = (message, type = 'info') => {
            if (toastTimeout) clearTimeout(toastTimeout);
            dom.toast.textContent = message;
            dom.toast.className = `d-dart-toast ${type} show`;
            toastTimeout = setTimeout(() => { dom.toast.classList.remove('show'); }, CONFIG.UI.TOAST_DURATION);
        };

        return {
            init() {
                injectStyles(); createDOM();
                debouncedAnalyze = debounce(startAnalysis, CONFIG.UI.SCROLL_DEBOUNCE, true);
                setupEventListeners(); setupStateSubscriptions();
                TokenManager.init(); ModeManager.init();
                this.updateSOWIndicator();
                updateModeIndicator();
                if (CONFIG.START_MINIMIZED) { dom.container.classList.add('minimized'); AppState.set('isMinimized', true); }
                Logger.info('UI initialized');
                Telemetry.track(TelemetryEventType.APP_INIT, { isOnSMC: Helpers.isOnSMC(), startMinimized: CONFIG.START_MINIMIZED, mode: ModeManager.getMode() });
            },

            updateTokenIndicator() { updateMinimizedState(); },
            updateSOWIndicator() { updateMinimizedState(); checkSOWErrorDisplay(); },

            showProgress(steps) {
                if (AppState.get('isSettingsOpen')) closeSettings();
                progressDOMRefs = null;
                const stepsHtml = steps.map(step => `
                    <div class="d-dart-step pending" id="d-dart-step-${step.id}" role="listitem">
                        <span class="d-dart-step-icon" aria-hidden="true">${step.icon}</span>
                        <span class="d-dart-step-text">${Helpers.escapeHtml(step.text)}</span>
                        <span class="d-dart-step-status" id="d-dart-step-status-${step.id}"></span>
                    </div>
                `).join('');
                dom.results.innerHTML = `<div class="d-dart-progress"><div class="d-dart-progress-title">🔄 Processing Order...</div><div class="d-dart-progress-steps" role="list">${stepsHtml}</div></div>`;
            },

            updateProgressStep(stepId, status, statusText = '') {
                const stepEl = document.getElementById(`d-dart-step-${stepId}`);
                const statusEl = document.getElementById(`d-dart-step-status-${stepId}`);
                if (stepEl) stepEl.className = `d-dart-step ${status}`;
                if (statusEl && statusText) statusEl.textContent = statusText;
            },

            // Fix #11: cache progress DOM refs
            showEnhancedBatchProgress(totalOrders, totalChunks) {
                if (AppState.get('isSettingsOpen')) closeSettings();
                dom.results.innerHTML = HTMLGenerator.enhancedBatchProgress(totalOrders, totalChunks);
                progressDOMRefs = {
                    processed: document.getElementById('d-dart-stat-processed'),
                    success: document.getElementById('d-dart-stat-success'),
                    failed: document.getElementById('d-dart-stat-failed'),
                    remaining: document.getElementById('d-dart-stat-remaining'),
                    bar: document.getElementById('d-dart-progress-bar'),
                    status: document.getElementById('d-dart-progress-status'),
                    eta: document.getElementById('d-dart-progress-eta'),
                    chunk: document.getElementById('d-dart-chunk-info'),
                    token: document.getElementById('d-dart-token-status')
                };
            },

            // Fix #11: use cached refs
            updateEnhancedBatchProgress(data) {
                const { processed, success, failed, total, startTime } = data;
                const remaining = total - processed;
                const percent = Math.round((processed / total) * 100);
                const refs = progressDOMRefs;
                if (!refs) return;
                if (refs.processed) refs.processed.textContent = processed;
                if (refs.success) refs.success.textContent = success;
                if (refs.failed) refs.failed.textContent = failed;
                if (refs.remaining) refs.remaining.textContent = remaining;
                if (refs.bar) { refs.bar.style.width = `${percent}%`; refs.bar.setAttribute('aria-valuenow', percent); }
                const currentChunk = AppState.get('currentChunk') + 1;
                const totalChunks = AppState.get('totalChunks');
                if (refs.chunk) refs.chunk.textContent = `${currentChunk}/${totalChunks}`;
                if (processed >= 5 && startTime) {
                    const elapsed = Date.now() - startTime; const avgTimePerOrder = elapsed / processed;
                    const etaMs = remaining * avgTimePerOrder;
                    if (refs.eta) refs.eta.textContent = `ETA: ${Helpers.formatETA(etaMs)}`;
                }
                if (refs.token) {
                    const remainingSec = TokenManager.getRemainingSeconds();
                    refs.token.textContent = remainingSec > 0 ? `${remainingSec}s` : 'Expired';
                    refs.token.className = remainingSec > 30 ? 'token-ok' : remainingSec > 0 ? 'token-warning' : 'token-error';
                }
            },

            updateBatchStatus(status) {
                const refs = progressDOMRefs;
                if (refs?.status) refs.status.textContent = status;
                else { const statusEl = document.getElementById('d-dart-progress-status'); if (statusEl) statusEl.textContent = status; }
            },

            showBatchComplete(reportData) {
                if (AppState.get('isSettingsOpen')) closeSettings();
                progressDOMRefs = null;
                dom.results.innerHTML = HTMLGenerator.batchReportTable(reportData);
            },

            showProcessingError(message) {
                if (AppState.get('isSettingsOpen')) closeSettings();
                progressDOMRefs = null;
                dom.results.innerHTML = `<div class="d-dart-error" role="alert"><div class="d-dart-error-icon" aria-hidden="true">❌</div><div class="d-dart-error-title">Processing Failed</div><div class="d-dart-error-message">${Helpers.escapeHtml(message)}</div></div>`;
            },

            displaySingleOrderResults(orderData) {
                if (AppState.get('isSettingsOpen')) closeSettings();
                progressDOMRefs = null;
                const analysisResults = orderData?.analysisResults || [];
                const smcExecutionData = orderData?.smcExecutionData;
                const stops = orderData?.viewData?.stops || [];

                let html = '<div class="d-dart-results">';
                if (ModeManager.isReviewMode()) html += `<div class="d-dart-review-banner">${Messages.INFO.REVIEW_MODE_BANNER}</div>`;
                html += HTMLGenerator.detentionSummaryBanner(orderData);
                html += HTMLGenerator.shipperCard(orderData);
                html += '<div class="d-dart-section-title">Stop Analysis</div>';
                for (let i = 0; i < analysisResults.length; i++) html += HTMLGenerator.stopCard(stops[i], analysisResults[i], orderData?.sowConfig, smcExecutionData);
                html += '</div>';
                dom.results.innerHTML = html;
            },

            showToast,

            // Fix #7: remove keyboard handler on destroy
            destroy() {
                const moveDrag = boundHandlers.get('moveDrag'); const endDrag = boundHandlers.get('endDrag');
                if (moveDrag) { document.removeEventListener('mousemove', moveDrag); document.removeEventListener('touchmove', moveDrag); }
                if (endDrag) { document.removeEventListener('mouseup', endDrag); document.removeEventListener('touchend', endDrag); }
                boundHandlers.clear(); AppState.clearListeners();
                if (toastTimeout) { clearTimeout(toastTimeout); toastTimeout = null; }
                // Fix #7: remove keyboard shortcut listener
                if (keyboardShortcutHandler) { document.removeEventListener('keydown', keyboardShortcutHandler); keyboardShortcutHandler = null; }
                // Fix #3: remove settings escape listener
                if (settingsEscapeHandler) { document.removeEventListener('keydown', settingsEscapeHandler); settingsEscapeHandler = null; }
                if (AppState.get('isSettingsOpen')) closeSettings();
                progressDOMRefs = null;
                if (dom?.container) dom.container.remove();
                if (dom?.toast) dom.toast.remove();
                ApprovalPopup.cleanup(); TokenManager.cleanup(); CacheManager.cleanup();
                ProgressManager.clear(); Telemetry.cleanup();
                Logger.info('UI destroyed and cleaned up');
            }
        };
    })();

    /* ═══════════════════════════════════════════════════════════════════════════
     * SECTION 33: APPLICATION INITIALIZATION
     * ═══════════════════════════════════════════════════════════════════════════ */

    const App = {
        async init() {
            window.addEventListener('unhandledrejection', (event) => {
                const message = event.reason?.message || String(event.reason || 'Unknown rejection');
                Logger.error('Unhandled promise rejection', { message, stack: event.reason?.stack, type: event.reason?.type });
            });

            Logger.info(`=== D-DART v${CONFIG.VERSION} ${CONFIG.APP_SUBTITLE} Starting ===`);
            Logger.info(`Page: ${window.location.href}`);
            Logger.info(`On SMC: ${Helpers.isOnSMC()}`);
            Logger.info(`Max Batch Size: ${CONFIG.BATCH.MAX_ORDERS_PER_SESSION}`);

            try {
                UIController.init();
                Logger.info('Loading SOW configuration from SharePoint...');
                await SOWConfigManager.init();
                Logger.info(`Health Check: ${HealthCheck.getSummary()}`);
                Logger.info(`Operation Mode: ${ModeManager.getMode()}`);
                Logger.info(`=== D-DART v${CONFIG.VERSION} ${CONFIG.APP_SUBTITLE} Ready ===`);
                Telemetry.track(TelemetryEventType.APP_INIT, { success: true, sowLoaded: SOWConfigManager.isLoaded(), shipperCount: SOWConfigManager.getShipperCount(), mode: ModeManager.getMode() });
            } catch (error) {
                Logger.error('Initialization failed', error.message);
                console.error('D-DART initialization failed:', error);
                Telemetry.track(TelemetryEventType.APP_ERROR, { phase: 'init', error: error.message });
            }
        },

        getInfo() {
            return {
                name: CONFIG.APP_NAME, version: CONFIG.VERSION, edition: CONFIG.APP_SUBTITLE, author: CONFIG.AUTHOR,
                operationMode: ModeManager.getMode(), maxBatchSize: CONFIG.BATCH.MAX_ORDERS_PER_SESSION,
                state: AppState.getSnapshot(), tokenStatus: TokenManager.getStatus(), sowStatus: SOWConfigManager.getStatus(),
                cacheStats: CacheManager.getStats(), batchState: BatchProcessor.getState(), health: HealthCheck.check(),
                performance: PerformanceMonitor.getMetrics(), telemetry: Telemetry.getMetrics()
            };
        },

        destroy() { UIController.destroy(); SOWConfigManager.clear(); Logger.info('Application destroyed'); }
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => App.init());
    } else {
        App.init();
    }

    if (CONFIG.DEBUG_ENABLED) {
        window.DDART = {
            App, AppState, Logger, TokenManager, SOWConfigManager, CacheManager, ProgressManager,
            BatchProcessor, HealthCheck, PerformanceMonitor, Telemetry, DetentionAnalyzer, ModeManager, LaneValidator,
            CONFIG, version: CONFIG.VERSION, edition: CONFIG.APP_SUBTITLE,
            test: {
                analyzeOrder: (orderId) => BatchProcessor.processBatch([orderId]),
                analyzeBatch: (orderIds) => BatchProcessor.processBatch(orderIds),
                getState: () => AppState.getSnapshot(),
                getHealth: () => HealthCheck.check(),
                isHealthy: () => HealthCheck.isHealthy(),
                clearCache: () => CacheManager.clear(),
                clearProgress: () => ProgressManager.clear(),
                refreshToken: () => TokenManager.ensure(),
                refreshSOW: () => SOWConfigManager.fetch(),
                getSOWConfig: (shipper) => SOWConfigManager.getConfig(shipper),
                getAllShippers: () => SOWConfigManager.getShipperNames(),
                getAllShippersData: () => SOWConfigManager.getAllShippersData(),
                getSOWStats: () => SOWConfigManager.getStatistics(),
                showToast: (msg, type) => UIController.showToast(msg, type),
                generateReport: () => Logger.generateReport(),
                pauseBatch: () => BatchProcessor.pause(),
                resumeBatch: () => BatchProcessor.resume(),
                cancelBatch: () => BatchProcessor.cancel(),
                getInfo: () => App.getInfo(),
                previewCharge: (params) => DetentionAnalyzer.previewCharge(params),
                getPerformance: () => PerformanceMonitor.getMetrics(),
                getTelemetry: () => Telemetry.getMetrics(),
                getMode: () => ModeManager.getMode(),
                toggleMode: () => ModeManager.toggle(),
                validateLane: (viewData, vr) => LaneValidator.validate(viewData, vr),
                resetCircuitBreakers: () => { Object.values(circuitBreakers).forEach(cb => cb.reset()); Logger.info('All circuit breakers reset'); }
            }
        };
    }

})();