// ==UserScript==
// @name         FIF智慧教学平台全自动刷课助手 (icourse.fifedu.com)
// @name:zh-CN   FIF智慧教学平台全自动刷课助手 (icourse.fifedu.com)
// @name:en      FIF Smart Teaching Platform Auto Study Helper
// @namespace    https://github.com/XiaoMaxer/course-learning-toolkit
// @version      1.9.41
// @description  适用于 FIF 智慧教学平台 (icourse.fifedu.com) 全站课程自动学习：支持学习空间首页课程识别、全自动巡航章节、视频自动倍速/秒过（含损坏视频容错）、PPT自动翻页完成、Word/PDF文档自动置底完成、自适应练习 AI 自动答题并交卷、课程「测评评估掌握度」自动刷题、自动跳过单元检测及防死循环保护；修复首次进入测评入口不渲染、测评必须手动选择学习目标才能启动的问题；新增测评「循环测试」开关与目标熟练度指标（选填 0-100 整数，实时校验），未达目标自动再测一轮，达标/手动停止/异常即终止；新增「无限循环刷熟练度」，忽略目标持续测评直到手动停止。
// @description:zh-CN 适用于 FIF 智慧教学平台 (icourse.fifedu.com) 全站课程自动学习：支持学习空间首页课程识别、全自动巡航章节、视频自动倍速/秒过（含损坏视频容错）、PPT自动翻页完成、Word/PDF文档自动置底完成、自适应练习 AI 自动答题并交卷、课程「测评评估掌握度」自动刷题、自动跳过单元检测及防死循环保护；修复首次进入测评入口不渲染、测评必须手动选择学习目标才能启动的问题；新增测评「循环测试」开关与目标熟练度指标（选填 0-100 整数，实时校验），未达目标自动再测一轮，达标/手动停止/异常即终止；新增「无限循环刷熟练度」，忽略目标持续测评直到手动停止。
// @description:en Automated course study helper for FIF Smart Teaching Platform (icourse.fifedu.com). Supports dashboard course detection, video, PPT, PDF/Word, AI-powered adaptive practice and course evaluation auto-completion.
// @author       Antigravity
// @match        https://icourse.fifedu.com/*
// @match        https://resc.fifedu.com/iplat-resc/pdfv-resource/web/viewer.html*
// @match        https://*.ow365.cn/*
// @icon         https://icourse.fifedu.com/favicon.ico
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_addStyle
// @grant        GM_registerMenuCommand
// @grant        unsafeWindow
// @grant        GM_xmlhttpRequest
// @grant        GM_setClipboard
// @connect      api.deepseek.com
// @connect      127.0.0.1
// @connect      localhost
// @run-at       document-idle
// @license      MIT
// ==/UserScript==

(function () {
    'use strict';

    const pageWindow = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    const host = location.hostname;

    // 平台用 window.confirm 提示「正在答题中」，原生对话框会冻结整页 JS（脚本与计时器全部停摆），
    // 且无法通过 DOM 点击，只能从源头拦截：按当前是否仍有作答题目决定返回「取消(false)」还是「确定(true)」
    if (host.includes('icourse.fifedu.com')) {
        const nativeConfirm = typeof pageWindow.confirm === 'function' ? pageWindow.confirm.bind(pageWindow) : null;
        const nativeAlert = typeof pageWindow.alert === 'function' ? pageWindow.alert.bind(pageWindow) : null;
        const LEAVE_CONFIRM_RE = /正在答题中|暂时保留|作答进度/;

        // true = 允许离开；false = 留在当前页继续作答
        const allowLeavingPractice = () => {
            try {
                if (state.leavingIntentionally) return true;
                const q = readPracticeQuestionDom();
                if (!q || q.over) return true;
                const now = Date.now();
                if (!state.confirmBlockedAt || now - state.confirmBlockedAt > 60000) state.confirmBlockedCount = 0;
                state.confirmBlockedAt = now;
                state.confirmBlockedCount++;
                if (state.confirmBlockedCount > 3) {
                    log('多次触发「正在答题中」离开确认，判定当前练习无法继续，本次运行跳过该练习', 'warn');
                    markExerciseFailed(readCurrentSectionName(), state.activeExerciseName || '', '反复触发离开确认');
                    return true;
                }
                log('已拦截平台「正在答题中」离开确认，留在当前页继续作答', 'warn');
                return false;
            } catch (e) {
                return true;
            }
        };

        pageWindow.confirm = function (msg) {
            try {
                if (LEAVE_CONFIRM_RE.test(String(msg || ''))) return allowLeavingPractice();
            } catch (e) { /* ignore */ }
            return nativeConfirm ? nativeConfirm(msg) : true;
        };

        pageWindow.alert = function (msg) {
            try {
                if (LEAVE_CONFIRM_RE.test(String(msg || ''))) {
                    log(`已拦截平台提示：${String(msg || '').replace(/\s+/g, ' ').slice(0, 80)}`, 'warn');
                    return;
                }
            } catch (e) { /* ignore */ }
            return nativeAlert ? nativeAlert(msg) : undefined;
        };
    }


    // =========================================================================
    // 1. 子框架辅助模块：跨域 PDF.js 阅读器 & ow365 PPT 预览器自动置底/翻页
    // =========================================================================
    if (host.includes('resc.fifedu.com')) {
        // PDF.js / Word 预览器子页面自动翻至末页并通知父窗口
        const runPdfHelper = () => {
            const interval = setInterval(() => {
                try {
                    const app = pageWindow.PDFViewerApplication;
                    if (app && app.pagesCount && app.pagesCount > 0) {
                        const total = app.pagesCount;
                        if (app.page < total) {
                            app.page = total;
                        }
                        const lastBtn = document.querySelector('#lastPage');
                        if (lastBtn) lastBtn.click();

                        const vc = document.querySelector('#viewerContainer');
                        if (vc) {
                            vc.scrollTop = vc.scrollHeight;
                            vc.dispatchEvent(new Event('scroll'));
                        }
                        pageWindow.parent.postMessage({
                            eventName: 'ready',
                            totalPage: total,
                            loading: false
                        }, '*');
                        pageWindow.parent.postMessage({
                            eventName: 'currentPage',
                            currentPage: total,
                            totalPage: total
                        }, '*');
                    }
                } catch (e) {
                    // ignore cross-frame errors
                }
            }, 2000);
            setTimeout(() => clearInterval(interval), 120000);
        };
        runPdfHelper();
        return;
    }

    if (host.includes('ow365.cn')) {
        // ow365 PPT 跨域预览器自动翻页至末页
        const runOw365Helper = () => {
            let lastPage = 0, lastPageTicks = 0;
            const clickNextHard = () => {
                // 常规下一页按钮
                const cands = Array.from(document.querySelectorAll(
                    '#pageNext, .pageNext, .next, .btnNext, [class*="next"], [title*="下一页"], [title*="下页"], [aria-label*="下一页"]'
                ));
                cands.forEach(el => {
                    try {
                        if (el.offsetWidth > 0 || el.offsetHeight > 0) el.click();
                    } catch (e) { /* ignore */ }
                });
                // 键盘兜底（部分皮肤只响应键盘）
                const targets = [document.activeElement, document.body, document.documentElement].filter(Boolean);
                ['ArrowRight', 'PageDown', ' '].forEach(key => {
                    targets.forEach(t => {
                        try {
                            t.dispatchEvent(new KeyboardEvent('keydown', {
                                key, code: key === ' ' ? 'Space' : key, bubbles: true
                            }));
                        } catch (e) { /* ignore */ }
                    });
                });
                // 点击幻灯片右侧区域兜底（ow365 支持点击右半页翻页）
                try {
                    const canvas = document.querySelector('canvas, .slide, .slideBox, #slideContainer') || document.body;
                    const r = canvas.getBoundingClientRect();
                    if (r.width > 100 && r.height > 100) {
                        const el = document.elementFromPoint(r.left + r.width * 0.85, r.top + r.height * 0.5);
                        if (el) el.click();
                    }
                } catch (e) { /* ignore */ }
            };
            const interval = setInterval(() => {
                try {
                    const allEls = Array.from(document.querySelectorAll('*'));
                    const pageEl = allEls.find(el => /^\d+\s*\/\s*\d+$/.test((el.innerText || '').trim()));
                    if (pageEl) {
                        const parts = pageEl.innerText.trim().split('/').map(s => parseInt(s.trim(), 10));
                        if (parts.length === 2 && !isNaN(parts[0]) && !isNaN(parts[1])) {
                            const [cur, total] = parts;
                            pageWindow.parent.postMessage({
                                currentPageNum: cur,
                                totalPageNum: total,
                                data: { currentPageNum: cur, totalPageNum: total }
                            }, '*');
                            if (cur >= total) {
                                clearInterval(interval);
                                return;
                            }
                        }
                    }
                    const nextBtn = document.querySelector('#pageNext, .pageNext, [title*="下一页"], [title*="下一张"]');
                    if (nextBtn) nextBtn.click();

                    // 页码连续多轮不推进：启用更强的翻页兜底（按钮/键盘/点击右侧区域）
                    const curPage = pageEl ? parseInt((pageEl.innerText || '0/1').split('/')[0], 10) || 0 : 0;
                    if (curPage > 0 && curPage === lastPage) {
                        lastPageTicks++;
                        if (lastPageTicks >= 6) clickNextHard();
                    } else {
                        lastPage = curPage;
                        lastPageTicks = 0;
                    }
                } catch (e) {
                    // ignore
                }
            }, 450);
            setTimeout(() => clearInterval(interval), 120000);
        };
        runOw365Helper();
        return;
    }

    // 仅在顶层窗口运行主控面板与状态机
    if ( window.top !== window.self ) return;

    // 单实例保护：同一页面只允许一个助手实例运行，避免重复安装时两个实例互相刷新面板（按钮闪烁）
    if (pageWindow.__fifHelperInstanceId) {
        console.warn('[FIF助手] 检测到脚本已被另一个实例加载，本次实例不运行。请在 Tampermonkey 中只保留一个脚本。');
        return;
    }
    pageWindow.__fifHelperInstanceId = 'fif-' + Date.now() + '-' + Math.random().toString(36).slice(2);

    // 接收子框架（ow365 / PDF.js）通过 postMessage 上报的页码，用于真实翻页进度校验
    const pptFrameState = { current: 0, total: 0, at: 0 };
    try {
        window.addEventListener('message', (ev) => {
            const d = ev && ev.data;
            if (!d || typeof d !== 'object') return;
            // 忽略脚本自己发给平台的模拟翻页消息，否则会把自己的假页码当成真实页码
            if (d.__fifSelf && d.__fifSelf === pageWindow.__fifHelperInstanceId) return;
            const cur = Number(d.currentPageNum ?? d.data?.currentPageNum ?? d.currentPage);
            const tot = Number(d.totalPageNum ?? d.data?.totalPageNum ?? d.totalPage);
            if (Number.isFinite(tot) && tot > 0) pptFrameState.total = tot;
            if (Number.isFinite(cur) && cur > 0) {
                // 页码不能超过总页数，避免残留状态串到下一个课件
                pptFrameState.current = (pptFrameState.total > 0 && cur > pptFrameState.total) ? pptFrameState.total : cur;
                pptFrameState.at = Date.now();
            }
        });
    } catch (e) { /* ignore */ }

    // =========================================================================
    // 2. 存储与全局配置管理 (兼容无 GM_* 环境的降级存储)
    // =========================================================================
    const Storage = {
        get(key, defaultVal) {
            try {
                if (typeof GM_getValue === 'function') {
                    const val = GM_getValue(key, undefined);
                    return val !== undefined ? val : defaultVal;
                }
                const raw = localStorage.getItem('fif_helper_' + key);
                return raw !== null ? JSON.parse(raw) : defaultVal;
            } catch (e) {
                return defaultVal;
            }
        },
        set(key, val) {
            try {
                if (typeof GM_setValue === 'function') {
                    GM_setValue(key, val);
                }
                localStorage.setItem('fif_helper_' + key, JSON.stringify(val));
            } catch (e) {
                // ignore
            }
        }
    };

    // AI 答题默认提示词（可在面板中修改）
    const DEFAULT_AI_PROMPT = '你是一名严谨的学科老师，正在为学生完成在线练习。只根据题干和选项作答，不要解释，不要输出多余文字。\n必须严格输出 JSON：{"answers":[["A"],["B","C"]]}。\n外层数组顺序与题目顺序一致；每个子数组是该题选择的选项字母。\n单选题和判断题只放一个字母；多选题放多个字母并按字母升序。\n判断题中 A 表示“正确”，B 表示“错误”。\n即使不确定也要给出最可能的答案，禁止留空。';

    const DEFAULT_CONFIG = {
        running: false,            // 是否开启自动刷课
        fastMode: true,            // 极速完成模式 (结合 Vue 进度上报与快进)
        videoRate: 4.0,            // 视频播放倍速 (1 ~ 16)
        autoExercise: true,        // 自动完成自适应练习
        skipExams: true,           // 自动跳过单元检测/测验
        autoNextCourse: false,     // 课程刷完后是否自动返回首页切换下一门课
        minimized: false,          // 面板是否最小化
        // ---- AI 自动答题配置 ----
        aiEnabled: true,           // 是否启用 AI 自动答题
        aiApiKey: '',               // DeepSeek API Key（请在脚本面板中填写）
        aiBaseUrl: 'https://api.deepseek.com/chat/completions',
        aiModel: 'deepseek-chat',
        aiPrompt: DEFAULT_AI_PROMPT,
        aiMinDelayMs: 2500,        // 每题提交前的最短模拟思考时间
        aiMaxDelayMs: 6000,        // 每题提交前的最长模拟思考时间
        aiMaxQuestions: 50,        // 单次练习最多自动作答题目数
        aiAutoSubmit: true,        // 答完后自动点击「显示报告/交卷」
        // ---- 熟练度测评配置 ----
        evalLoop: false,           // 循环测试：本次熟练度低于目标值时自动再测一轮（默认关闭）
        evalTargetMastery: '',     // 目标熟练度指标（选填，0-100 的整数）
        evalInfinite: false,       // 无限循环刷熟练度：忽略目标值，持续测评直到手动停止（默认关闭）
        fullAutoMastery: false,    // 私有：一键刷全部课程并完成掌握度指标（运行中标记，跨页面持久化）
        fullAutoChapterThreshold: 95, // 私有：一键达标阶段 A 的章节完成阈值（0-100 整数）
        fullAutoPhase: 'chapters'  // 私有：一键达标当前阶段：chapters（刷章节）/ mastery（刷掌握度）
    };

    const state = {
        config: {
            running: Storage.get('running', DEFAULT_CONFIG.running),
            fastMode: Storage.get('fastMode', DEFAULT_CONFIG.fastMode),
            videoRate: Number(Storage.get('videoRate', DEFAULT_CONFIG.videoRate)) || 4.0,
            autoExercise: Storage.get('autoExercise', DEFAULT_CONFIG.autoExercise),
            skipExams: Storage.get('skipExams', DEFAULT_CONFIG.skipExams),
            autoNextCourse: Storage.get('autoNextCourse', DEFAULT_CONFIG.autoNextCourse),
            minimized: Storage.get('minimized', DEFAULT_CONFIG.minimized),
            aiEnabled: Storage.get('aiEnabled', DEFAULT_CONFIG.aiEnabled),
            aiApiKey: Storage.get('aiApiKey', DEFAULT_CONFIG.aiApiKey),
            aiBaseUrl: Storage.get('aiBaseUrl', DEFAULT_CONFIG.aiBaseUrl),
            aiModel: Storage.get('aiModel', DEFAULT_CONFIG.aiModel),
            aiPrompt: Storage.get('aiPrompt', DEFAULT_CONFIG.aiPrompt),
            aiMinDelayMs: Number(Storage.get('aiMinDelayMs', DEFAULT_CONFIG.aiMinDelayMs)) || DEFAULT_CONFIG.aiMinDelayMs,
            aiMaxDelayMs: Number(Storage.get('aiMaxDelayMs', DEFAULT_CONFIG.aiMaxDelayMs)) || DEFAULT_CONFIG.aiMaxDelayMs,
            aiMaxQuestions: Number(Storage.get('aiMaxQuestions', DEFAULT_CONFIG.aiMaxQuestions)) || DEFAULT_CONFIG.aiMaxQuestions,
            aiAutoSubmit: Storage.get('aiAutoSubmit', DEFAULT_CONFIG.aiAutoSubmit),
            evalLoop: Storage.get('evalLoop', DEFAULT_CONFIG.evalLoop),
            evalTargetMastery: Storage.get('evalTargetMastery', DEFAULT_CONFIG.evalTargetMastery),
            evalInfinite: Storage.get('evalInfinite', DEFAULT_CONFIG.evalInfinite),
            fullAutoMastery: Storage.get('fullAutoMastery', DEFAULT_CONFIG.fullAutoMastery),
            fullAutoChapterThreshold: clampChapterThreshold(Storage.get('fullAutoChapterThreshold', DEFAULT_CONFIG.fullAutoChapterThreshold)),
            fullAutoPhase: Storage.get('fullAutoPhase', DEFAULT_CONFIG.fullAutoPhase)
        },
        skipSections: new Set(Storage.get('skipSections', [])),
        // 平台明确判定「已进行过该练习 / 知识点已掌握 / 无新题」的活动（小节::活动），可放心视为已完成
        doneExercises: new Set(Storage.get('doneExercises', [])),
        // 本次运行进不去的练习：只做会话级跳过，刷新后重试，避免误判整节完成
        sessionFailedExercises: new Set(),
        exerciseFailCount: new Map(),
        // 本地判定「已处理完」的小节（剩余未完成项均为平台确认已完成），避免目录页反复进入
        handledSections: new Set(Storage.get('handledSections', [])),
        sessionSkipSections: new Set(),
        // 导航落错小节（点了 A 实际进 B）的连续失败计数，仅本次运行有效
        navFailCount: new Map(),
        // 平台侧活动全部完成但目录状态迟迟不同步的小节，记录进入次数以限制反复横跳
        platformDoneVisits: new Map(),
        sectionFailCount: new Map(Object.entries(Storage.get('sectionFailCounts', {})).map(([k, v]) => [k, Number(v) || 0])),
        skipReasonLogged: new Set(),
        completedCourses: new Set(Storage.get('completedCourses', [])),
        masteryMetCourses: new Set(Storage.get('masteryMetCourses', [])),
        chapterDoneCourses: new Set(Storage.get('chapterDoneCourses', [])),
        chapterRetriedCourses: new Set(Storage.get('chapterRetriedCourses', [])),
        currentChapterKey: '',
        processedInVisit: new Set(),
        lastIndexUrl: Storage.get('lastIndexUrl', ''),
        busy: false,
        evaluationRunning: false,
        evalStopRequested: false,
        stuckCounter: 0,
        lastProgressAt: Date.now(),
        watchdogPhase: 0,
        lastModalSnapshot: '',
        leavingIntentionally: false,
        confirmBlockedCount: 0,
        confirmBlockedAt: 0,
        activeExerciseName: '',
        // 刚从首页点卡片进课程的窗口期（毫秒时间戳）：此期间禁止 Vue Router 落到 pattern/chapter，强制改到 /index
        routerRedirectUntil: 0,
        // 目录页扫描期间钉住 /index 的截止时间：平台会自动跳回"上次学习的小节"，此期间一律反弹回目录页
        indexPinUntil: 0,
        // 标记：目录页补扫期间，强制忽略平台侧"已完成"状态，从头重新扫描
        forceRescanIndex: false,
        // 目录页点小节前设置的目标小节，用于学习页校验是否真进到了那一节
        targetSectionName: '',
        targetSectionChapterId: '',
        statusText: '就绪，点击「开始刷课」启动自动化',
        logs: []
    };

    function saveConfig() {
        Object.keys(DEFAULT_CONFIG).forEach(k => Storage.set(k, state.config[k]));
    }

    function saveSkipSections() {
        Storage.set('skipSections', Array.from(state.skipSections));
    }

    // 判断元素 class 中是否存在指定的状态类名，避免 [class*="isfinish"] 这类模糊匹配误判
    function hasStatusClassToken(root, tokenRe) {
        if (!root) return false;
        return Array.from(root.querySelectorAll('[class]')).some(el => {
            const cls = typeof el.className === 'string' ? el.className : '';
            return cls.split(/\s+/).some(c => c && tokenRe.test(c));
        });
    }

    // 同一小节连续失败 3 次才永久跳过；前 2 次仅本次运行跳过，重新加载页面后会再次尝试
    function markSectionSkipped(name) {
        if (!name) return;
        const n = (state.sectionFailCount.get(name) || 0) + 1;
        state.sectionFailCount.set(name, n);
        saveSectionFailCounts();
        if (n >= 3) {
            state.skipSections.add(name);
            saveSkipSections();
            log(`小节【${name}】已连续 ${n} 次未能完成，加入永久跳过列表（可在面板「清空跳过缓存」恢复）`, 'warn');
        } else {
            state.sessionSkipSections.add(name);
            log(`小节【${name}】本次运行暂时跳过（第 ${n}/3 次失败），重新加载页面后会再次尝试`, 'warn');
        }
    }

    // 跳过原因只记录一次，避免每秒刷屏
    function logSkipOnce(name, reason) {
        const key = `${name}｜${reason}`;
        if (state.skipReasonLogged.has(key)) return;
        state.skipReasonLogged.add(key);
        if (state.skipReasonLogged.size > 200) state.skipReasonLogged.clear();
        log(`跳过小节【${name}】：${reason}`, 'info');
    }

    function saveCompletedCourses() {
        Storage.set('completedCourses', Array.from(state.completedCourses));
    }

    function saveMasteryMetCourses() {
        Storage.set('masteryMetCourses', Array.from(state.masteryMetCourses));
    }

    function saveChapterDoneCourses() {
        Storage.set('chapterDoneCourses', Array.from(state.chapterDoneCourses));
    }

    function saveChapterRetriedCourses() {
        Storage.set('chapterRetriedCourses', Array.from(state.chapterRetriedCourses));
    }

    function saveDoneExercises() {
        Storage.set('doneExercises', Array.from(state.doneExercises));
    }

    function saveHandledSections() {
        Storage.set('handledSections', Array.from(state.handledSections));
    }

    // 小节已处理标记带课程 ID，避免不同课程同名小节互相影响
    function currentCourseId() {
        let id = new URLSearchParams(location.search).get('courseId');
        if (!id && state.lastIndexUrl) {
            try { id = new URL(state.lastIndexUrl).searchParams.get('courseId'); } catch (e) { id = ''; }
        }
        return id || '';
    }

    function sectionDoneKey(name) {
        return `${currentCourseId()}::${name || ''}`;
    }

    // 小节失败次数落盘，避免页面跳转/刷新后计数清零导致无限重试
    function saveSectionFailCounts() {
        const obj = {};
        state.sectionFailCount.forEach((v, k) => { obj[k] = v; });
        Storage.set('sectionFailCounts', obj);
    }

    function readCurrentSectionName() {
        const el = document.querySelector('.current-chapter-name, .chapter-name-title');
        return ((el && el.innerText) || '').split('\n')[0].trim();
    }

    // 平台判定「练习已做过 / 知识点已掌握 / 无新题」时的提示文案
    const PRACTICE_BLOCKED_PATTERN = /已经进行过(?:该|此)?练习|已进行过|已练习过|已完成过|练习已完成|已经完成(?:该|此)?练习|知识点已经掌握|已经掌握[，,]?可以选择其他活动|暂无可?练习|暂无新题|无新题|没有可练习|建议选择其他活动/;

    // 清除所有匹配 PRACTICE_BLOCKED_PATTERN 的 toast/弹窗 DOM，避免残留干扰后续活动判定
    function clearAllPracticeBlockedToasts() {
        try {
            const candidates = Array.from(document.querySelectorAll(
                '.el-message, .el-message__content, .el-message-group__wrapper, .el-notification, .el-notification__content, .el-message-box, .el-message-box__wrapper, .el-overlay, [class*="toast"], [class*="Toast"], .el-dialog, .el-dialog__wrapper'
            ));
            let removed = 0;
            for (const el of candidates) {
                const text = String(el.innerText || el.textContent || '').replace(/\s+/g, '');
                if (!text) continue;
                if (PRACTICE_BLOCKED_PATTERN.test(text)) {
                    // el-message 是 toast 容器，移除它就能彻底消除残留
                    const container = el.closest('.el-message, .el-message-box__wrapper, .el-notification, [class*="toast"], [class*="Toast"]') || el;
                    if (container && container.parentNode) {
                        container.parentNode.removeChild(container);
                        removed++;
                    } else {
                        // 兜底：清空文本 + 隐藏
                        container.innerText = '';
                        container.style.display = 'none';
                    }
                }
            }
            if (removed > 0) log(`已清理 ${removed} 条残留「已掌握/已做过」提示`, 'info');
        } catch (e) { /* ignore */ }
    }

    function readPracticeBlockedTip() {
        try {
            const candidates = [];
            document.querySelectorAll('.el-message, .el-message__content, .el-notification, .el-notification__content, [class*="toast"], [class*="Toast"], .el-message-box, .el-dialog')
                .forEach(el => { if (isVisibleEl(el)) candidates.push(el); });
            for (const el of candidates) {
                const text = String(el.innerText || el.textContent || '').replace(/\s+/g, '');
                const m = text.match(PRACTICE_BLOCKED_PATTERN);
                if (m) return m[0];
            }
            const scope = document.querySelector('.SubjectLib, .SubjectCover, .SubjectLayout, .SubjectCon');
            const scopeText = String((scope && scope.innerText) || '').replace(/\s+/g, '');
            const m2 = scopeText.match(PRACTICE_BLOCKED_PATTERN);
            if (m2) return m2[0];
        } catch (e) { /* ignore */ }
        return '';
    }

    // 平台把该练习标记为「练过 / 已完成」：页面只剩练习记录或掌握度变化，没有任何作答入口
    // 典型文案：近一个月掌握度变化 / 总练习1次 完成3题 正确率100% / 6小时之前 练习过
    function detectExerciseAlreadyDone() {
        try {
            if (readPracticeQuestionDom()) return '';
            const text = String(document.body.innerText || '').replace(/\s+/g, '');
            const hasRecord = /(分钟|小时|天|刚刚)之前练习过/.test(text) ||
                /近一个月掌握度/.test(text) ||
                (/总练习\d+次/.test(text) && /完成\d+题/.test(text));
            if (!hasRecord) return '';
            // 仍有重新作答入口时说明还能练，交给正常答题流程
            if (findClickableByText(/^(重新练[习题]|再练一次|再次练习|开始练习|继续练习|提升掌握度|提升熟练度|去练习)$/)) return '';
            return '练习已练过，无作答入口';
        } catch (e) { /* ignore */ }
        return '';
    }

    function exerciseKey(section, name) {
        return `${section || ''}::${name || ''}`;
    }

    // 平台明确表示该练习已做过/无需再做：落盘标记为完成，并让主循环顺延到本节下一个练习
    function markExerciseDone(section, name, reason) {
        if (!name) return;
        const key = exerciseKey(section, name);
        state.sessionFailedExercises.delete(key);
        if (state.doneExercises.has(key)) return;
        state.doneExercises.add(key);
        saveDoneExercises();
        // 清掉可能残留的"已掌握/已做过" toast，避免下一个活动被误判
        clearAllPracticeBlockedToasts();
        log(`练习【${name}】平台已判定无需再做（${reason}），标记完成并顺延到下一个练习`, 'success');
    }

    // 进不去/没题目属于弱失败：仅本次运行跳过，刷新后仍会重试，避免把整节误判为 100%
    function markExerciseFailed(section, name, reason) {
        if (!name) return;
        const key = exerciseKey(section, name);
        state.sessionFailedExercises.add(key);
        const n = (state.exerciseFailCount.get(key) || 0) + 1;
        state.exerciseFailCount.set(key, n);
        log(`练习【${name}】本次运行暂时跳过（${reason}，第 ${n} 次），顺延到下一个活动，刷新后会重试`, 'warn');
    }

    // 章节完成阈值归一化：0-100 整数，非法值回落到默认 95
    function clampChapterThreshold(raw) {
        const n = Number(raw);
        if (!Number.isFinite(n)) return DEFAULT_CONFIG.fullAutoChapterThreshold;
        const r = Math.round(n);
        if (r < 0) return 0;
        if (r > 100) return 100;
        return r;
    }

    // 私有功能：一键「全部课程 + 掌握度达标」是否运行中
    function isFullAutoMastery() {
        return !!state.config.fullAutoMastery;
    }

    // 结束一键达标任务：重置运行标记与阶段，保留已完成记录（可通过「清空跳过缓存」清除）
    function stopFullAutoMastery(reason) {
        state.config.fullAutoMastery = false;
        state.config.fullAutoPhase = 'chapters';
        state.config.running = false;
        state.evalStopRequested = true;
        saveConfig();
        updatePanelUI();
        if (reason) log(reason, 'success');
    }

    function log(msg, level = 'info') {
        const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
        const line = `[${time}] ${msg}`;
        console.log(`[FIF助手] ${line}`);
        state.logs.unshift({ time, msg, level });
        if (state.logs.length > 60) state.logs.pop();
        state.statusText = msg;
        state.lastProgressAt = Date.now();
        updatePanelUI();
    }

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    // =========================================================================
    // 3. Vue 3 实例与底层组件探测工具
    // =========================================================================
    function getVueApp() {
        const root = document.querySelector('#app, #microApp');
        return root ? root.__vue_app__ : null;
    }

    function getVueRouter() {
        const app = getVueApp();
        return app?.config?.globalProperties?.$router || null;
    }

    // 永久 hook Vue Router：在"刚从首页进课程"的窗口期内（state.routerRedirectUntil），
    // 任何 push/replace 到 /pattern/chapter 的请求都被重定向到 /index。
    // 这是为了防止平台从 sessionStorage 恢复"上次访问章节"而跳过中间小节。
    let __vueRouterRedirected = false;
    function installVueRouterRedirector() {
        if (__vueRouterRedirected) return;
        try {
            const router = getVueRouter();
            if (!router || typeof router.push !== 'function') return;
            const origPush = router.push.bind(router);
            const origReplace = router.replace ? router.replace.bind(router) : null;

            // null=不干预；'block'=扫描期间吞掉跳转（钉在目录页）；'redirect'=刚进课程窗口期改道目录页
            const classifyTarget = (target) => {
                const path = typeof target === 'string' ? target : target?.path || '';
                const name = typeof target === 'object' && target ? String(target.name || '') : '';
                if (!path.includes('/pattern/chapter') && !/pattern|chapter/i.test(name)) return null;
                if (Date.now() < state.indexPinUntil) return 'block';
                if (Date.now() < state.routerRedirectUntil) return 'redirect';
                return null;
            };

            // 扫描期间平台会定时把页面自动带回"上次学习的小节"。直接吞掉这次跳转
            // （不导航、不改道），既能让目录树渲染完，又不会触发"平台再跳→我们再改道"的紧循环。
            const shouldBlock = (target) => {
                if (classifyTarget(target) !== 'block') return false;
                const now = Date.now();
                if (!state.pinBlockLoggedAt || now - state.pinBlockLoggedAt > 5000) {
                    state.pinBlockLoggedAt = now;
                    log('已拦截平台自动跳回上次学习小节的跳转，停留在课程目录继续扫描', 'info');
                }
                return true;
            };

            const buildIndexTarget = (target) => {
                const query = {};
                try {
                    const sp = new URLSearchParams(location.search);
                    query.courseId = sp.get('courseId') || (typeof target === 'object' ? target?.query?.courseId : null);
                    query.classId = sp.get('classId') || (typeof target === 'object' ? target?.query?.classId : null);
                    query.termId = sp.get('termId') || (typeof target === 'object' ? target?.query?.termId : null);
                } catch (e) {}
                // 注意：Vue Router 的 base 是 /istp-learning-center/，这里必须用相对根路径 '/index'，
                // 否则会拼成 /istp-learning-center/istp-learning-center/index 这种重复前缀的坏 URL
                return { path: '/index', query };
            };

            router.push = function (target) {
                if (shouldBlock(target)) return Promise.resolve();
                if (classifyTarget(target) === 'redirect') {
                    const idx = buildIndexTarget(target);
                    log(`Vue Router push pattern/chapter → 重定向到 /index（刚进课程窗口期）`, 'info');
                    if (idx.query.courseId) {
                        const u = new URL(location.href);
                        u.pathname = '/istp-learning-center/index';
                        u.searchParams.set('courseId', idx.query.courseId);
                        if (idx.query.classId) u.searchParams.set('classId', idx.query.classId);
                        if (idx.query.termId) u.searchParams.set('termId', idx.query.termId);
                        Storage.set('lastIndexUrl', u.toString());
                    }
                    return origPush(idx);
                }
                return origPush(target);
            };

            if (origReplace) {
                router.replace = function (target) {
                    if (shouldBlock(target)) return Promise.resolve();
                    if (classifyTarget(target) === 'redirect') {
                        log('Vue Router replace pattern/chapter → 重定向到 /index（刚进课程窗口期）', 'info');
                        return origReplace(buildIndexTarget(target));
                    }
                    return origReplace(target);
                };
            }
            __vueRouterRedirected = true;
        } catch (e) { /* ignore */ }
    }

    function findVueComponent(selector, predicate) {
        const els = Array.from(document.querySelectorAll(selector));
        const visited = new Set();
        for (const el of els) {
            let curEl = el;
            while (curEl) {
                let comp = curEl.__vueParentComponent;
                while (comp && !visited.has(comp)) {
                    visited.add(comp);
                    if (!predicate || predicate(comp)) {
                        return comp;
                    }
                    comp = comp.parent;
                }
                curEl = curEl.parentElement;
            }
        }
        return null;
    }

    function hasEmitOption(comp, eventName) {
        const opts = comp?.emitsOptions || comp?.type?.emits;
        if (!opts) return false;
        if (Array.isArray(opts)) return opts.includes(eventName);
        if (typeof opts === 'object') return eventName in opts;
        return false;
    }

    /**
     * 通过 Vue 组件直接上报 100% 进度 (解决跨域 PDF / 单页文档 / 损坏视频无法打勾问题)
     */
    function triggerVueStudyProcess({ total = 100, end = 100, type = '275003' } = {}) {
        try {
            let triggered = false;

            // 方式 1: 直接调用 child preview 组件 (Pdf / Ppt / AudioVideo 等) 的 emit('studyProcess', ...)
            const childComp = findVueComponent('.wrapBox, .videoWrap, .pdfContainer, .pptWrap, .previewContainer', c => {
                return hasEmitOption(c, 'studyProcess') && !(c.props && c.props.resourceInfo && c.subTree?.el?.classList?.contains('previewContainer'));
            });
            if (childComp && typeof childComp.emit === 'function') {
                childComp.emit('studyProcess', { total, end, type });
                triggered = true;
            }

            // 方式 2: 通过 .previewContainer 向上 emit('studyProcess', payload)
            const previewComp = findVueComponent('.previewContainer', c => {
                return c.props && c.props.resourceInfo && typeof c.emit === 'function';
            });
            if (previewComp && previewComp.props?.resourceInfo) {
                const info = previewComp.props.resourceInfo;
                const courseId = new URLSearchParams(location.search).get('courseId') ||
                                 sessionStorage.getItem('courseId') || '';
                const payload = {
                    courseId: String(courseId).replace(/^"|"$/g, ''),
                    chapterId: info.chapterId,
                    activityId: info.activityId,
                    activityType: info.activityType || 1,
                    resourceId: info.resourceId,
                    resourceName: info.activityName,
                    resourceType: type,
                    videoTotalTime: total,
                    videoStudyTime: end
                };
                previewComp.emit('studyProcess', payload);
                triggered = true;
            }

            // 方式 3: 直接调用 .activity-list-warp 暴露的 upDateStudyProgress
            const actListComp = findVueComponent('.activity-list-warp, .chapter-info-right', c => {
                return c.exposed && typeof c.exposed.upDateStudyProgress === 'function';
            });
            if (actListComp && actListComp.exposed) {
                const list = actListComp.exposed.resoureList?.value || actListComp.exposed.resoureList || [];
                const activeItem = Array.isArray(list)
                    ? list.find(item => item.learnStatus !== 271003) || list[0]
                    : null;
                if (activeItem) {
                    actListComp.exposed.upDateStudyProgress({
                        activityId: activeItem.activityId,
                        resourceType: type,
                        videoTotalTime: total,
                        videoStudyTime: end
                    });
                    triggered = true;
                }
            }
            return triggered;
        } catch (e) {
            console.warn('[FIF助手] Vue 进度同步调用异常:', e);
        }
        return false;
    }

    // =========================================================================
    // 4. 通用弹窗与模态框处理
    // =========================================================================
    // 弹窗容器选择器（覆盖 Element Plus 的 MessageBox / Dialog 以及平台自定义 dialog 类）
    const MODAL_CONTAINER_SELECTORS = [
        '.el-message-box', '.el-dialog',
        '.el-message-box__wrapper', '.el-dialog__wrapper',
        '.el-overlay-message-box',
        '[class*="dialog"]', '[class*="Dialog"]'
    ].join(', ');
    // 允许自动点击的按钮文案（精确匹配，长度限制在 12 字内）
    const MODAL_ACCEPT_WORDS = ['退出', '仍要退出', '确定', '确认', '离开', '我知道了', '知道了', '好的', '好', '是', '仍要学习', '继续学习', '下一步', '关闭'];
    // 明确不点击的按钮文案（取消类操作保留给用户）
    const MODAL_REJECT_WORDS = ['取消', '暂不', '稍后', '再想想', '继续答题', '继续作答', '继续测评'];

    function modalTextOf(el) {
        return String((el && (el.innerText || el.textContent)) || '').replace(/\s+/g, ' ').trim();
    }

    // 收集当前可见的弹窗容器（去重，避免外层遮罩与内层弹窗重复计入）
    function visibleModals() {
        const raw = Array.from(document.querySelectorAll(MODAL_CONTAINER_SELECTORS));
        const out = [];
        raw.forEach(el => {
            if (!isVisibleEl(el)) return;
            if (out.some(parent => parent.contains(el))) return;
            const text = modalTextOf(el);
            const hasBtn = !!el.querySelector('button, .el-button, [role="button"]');
            if (!text && !hasBtn) return;
            out.push(el);
        });
        return out;
    }

    // 在弹窗内查找并点击匹配文案的按钮（兼容非 button 元素，如 div/span 伪按钮）
    function clickModalButton(scope, matcher) {
        const nodes = Array.from(scope.querySelectorAll(
            'button, .el-button, [role="button"], .el-message-box__btns span, .el-dialog__footer span'
        )).filter(isVisibleEl);
        for (const node of nodes) {
            const text = modalTextOf(node);
            if (!matcher(text)) continue;
            const clickable = node.closest('button, .el-button, [role="button"]') || node;
            if (clickable.disabled || clickable.classList.contains('is-disabled')) continue;
            fireHumanClick(clickable);
            return text.replace(/\s+/g, '') || '按钮';
        }
        return '';
    }

    async function handleModals() {
        const modals = visibleModals();
        if (!modals.length) return false;

        const isRejectText = (t) => MODAL_REJECT_WORDS.some(w => t.includes(w));
        const isAcceptText = (t) => {
            const compact = String(t || '').replace(/\s+/g, '');
            if (!compact || compact.length > 12) return false;
            if (isRejectText(compact)) return false;
            // 单字按钮（是/好）必须完全相等，避免误匹配「是否…」等文案
            return MODAL_ACCEPT_WORDS.some(w => (w.length === 1 ? compact === w : (compact === w || compact.includes(w))));
        };

        // 1) 「正在答题中，退出后为您暂时保留此次作答进度！」
        //    按场景区分：
        //    - 小节练习中且仍有未答题目：点「取消/继续答题」留在当前页，避免丢弃后面的题
        //    - 测评中 / 无活动题目：点「退出/确定」退出，继续后续流程（原行为）
        const exitModal = modals.find(m => /正在答题中|暂时保留|作答进度/.test(modalTextOf(m)));
        if (exitModal) {
            const inPractice = (() => {
                if (state.evaluationRunning) return false;
                if (document.querySelector('.CEDialog, .dialogBox, .HappyDrink')) return false;
                try { const q = readPracticeQuestionDom(); return !!(q && !q.over); } catch (e) { return false; }
            })();
            if (inPractice) {
                const stayed = clickModalButton(exitModal, (t) => {
                    const c = String(t || '').replace(/\s+/g, '');
                    if (!c) return false;
                    return /取消|继续答题|继续作答|留在|暂不|再想想/.test(c);
                });
                if (stayed) {
                    log(`检测到【正在答题中】弹窗，本节仍有题目，已点击【${stayed}】留在当前页继续作答`, 'warn');
                    state.lastProgressAt = Date.now();
                    await sleep(1200);
                    return true;
                }
            }
            const clicked = clickModalButton(exitModal, (t) => {
                const c = String(t || '').replace(/\s+/g, '');
                if (!c || isRejectText(c)) return false;
                return /退出|离开|确定|确认|^是$/.test(c);
            });
            if (clicked) {
                log(`检测到【正在答题中】弹窗，已点击【${clicked}】并保留作答进度`, 'warn');
                await sleep(1200);
                return true;
            }
            const snapshot = modalTextOf(exitModal).slice(0, 160);
            if (state.lastModalSnapshot !== snapshot) {
                state.lastModalSnapshot = snapshot;
                log(`检测到【正在答题中】弹窗但未找到退出按钮，内容：${snapshot}`, 'warn');
            }
        }

        // 2) 常规确认类弹窗：确定 / 确认 / 我知道了 / 仍要学习 / 下一步 ...
        for (const m of modals) {
            const clicked = clickModalButton(m, isAcceptText);
            if (clicked) {
                log(`自动点击弹窗按钮：【${clicked}】`, 'success');
                await sleep(1200);
                return true;
            }
        }

        // 3) 兜底：点击弹窗右上角关闭 X
        //    注意：closeBtn 必须严格限定到 Element UI 的关闭按钮类，
        //    不能用 [class*="close"] 这种模糊匹配（容易误命中非关闭元素）
        for (const m of modals) {
            const closeBtn = Array.from(m.querySelectorAll(
                '.el-message-box__headerbtn, .el-dialog__headerbtn, .el-message-box__close, .el-dialog__close, .el-message-box__headerbtn .el-icon-close, .el-dialog__headerbtn .el-icon-close'
            )).find(el => isVisibleEl(el));
            if (!closeBtn) continue;
            // 避免在同一弹窗上反复点：用元素唯一性标记
            const snap = closeBtn.outerHTML.slice(0, 80) + '|' + modalTextOf(m).slice(0, 40);
            if (state.lastModalSnapshot === snap) {
                // 已经点过同一个关闭按钮但弹窗没消失，放弃本次兜底
                log('已点击关闭按钮但弹窗未消失，跳过兜底关闭以免死循环', 'warn');
                return true;
            }
            state.lastModalSnapshot = snap;
            fireHumanClick(closeBtn);
            log('已点击弹窗右上角关闭按钮', 'warn');
            // 验证弹窗是否真消失；没消失就停止反复点击
            await sleep(800);
            const stillVisible = Array.from(document.querySelectorAll(MODAL_CONTAINER_SELECTORS))
                .some(el => isVisibleEl(el) && el.contains(m));
            if (stillVisible) {
                const stillSnap = modalTextOf(m).slice(0, 120);
                log(`点击关闭按钮后弹窗仍未消失：${stillSnap}，不再重试`, 'warn');
            }
            return true;
        }

        // 4) 存在弹窗但无法处理：记录文本快照留证（同一内容不重复刷屏）
        const snapshot = modals.map(modalTextOf).filter(Boolean).join(' | ').slice(0, 200);
        if (snapshot && state.lastModalSnapshot !== snapshot) {
            state.lastModalSnapshot = snapshot;
            log(`存在未识别弹窗，暂不处理：${snapshot}`, 'warn');
        }
        return false;
    }

    // 测评收尾清扫：关闭残留弹窗/遮罩，避免中途退出后弹窗卡住主循环
    async function cleanupEvaluationUi(reason = '') {
        for (let i = 0; i < 3; i++) {
            const handled = await handleModals();
            if (!handled) break;
            await sleep(400);
        }
        if (!visibleModals().length) {
            // 仅在没有可见弹窗时解除遮罩点击拦截，避免影响正常页面
            document.querySelectorAll('.el-overlay, .v-modal').forEach(el => {
                if (isVisibleEl(el)) el.style.pointerEvents = 'none';
            });
            document.body.classList.remove('el-popup-parent--hidden');
        }
        if (reason) log(`已清理测评残留界面：${reason}`, 'warn');
    }

    // =========================================================================
    // 5. 页面 A: 教学/学习空间首页 (/istp-manage/work/home)
    // =========================================================================
    function decorateWorkHomeCards() {
        const cards = Array.from(document.querySelectorAll('.learningCourses .cardToLearn'));
        cards.forEach((card) => {
            if (card.querySelector('.fif-helper-card-btn')) return;
            const nameEl = card.querySelector('.courseName');
            const courseName = nameEl ? nameEl.innerText.trim() : '未命名课程';

            const btn = document.createElement('button');
            btn.className = 'fif-helper-card-btn';
            btn.innerHTML = '🚀 自动刷此课';
            btn.title = `启动全自动刷课：${courseName}`;
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                e.preventDefault();
                state.config.running = true;
                saveConfig();
                log(`准备进入课程：【${courseName}】`, 'info');
                enterCourseCardInSameTab(card);
            });

            const infoBox = card.querySelector('.courseInfo') || card;
            infoBox.style.position = 'relative';
            infoBox.appendChild(btn);
        });
    }

    function enterCourseCardInSameTab(cardEl) {
        // 0. 先确保 Vue Router 守卫已安装（幂等），同时清理残留定时器
        installVueRouterRedirector();

        // 1. 彻底清 sessionStorage + localStorage 中平台存储的路由记忆
        //    平台 Vue Router 可能从任何地方恢复"上次访问章节"，所以全清
        try {
            const routeKeyRe = /chapter|pattern|route|visit|study|current|active|last|cache|keep|session|tab|history|prev|next|location|jump|navigate|cursor/i;
            ['sessionStorage', 'localStorage'].forEach(storeName => {
                try {
                    const store = window[storeName];
                    const keys = [];
                    for (let i = 0; i < store.length; i++) {
                        const k = store.key(i);
                        if (k && routeKeyRe.test(k)) keys.push(k);
                    }
                    keys.forEach(k => store.removeItem(k));
                } catch (e) { /* ignore per-store errors */ }
            });
            log('已清理平台路由记忆（sessionStorage + localStorage）', 'info');
        } catch (e) { /* ignore */ }

        // 2. 设置 Vue Router 窗口期（延长到 15 秒，等 Vue app 挂好 + card click 事件冒泡）+ 目录页补扫标记
        state.routerRedirectUntil = Date.now() + 15000;
        state.forceRescanIndex = true;

        // 3. 拦截 window.open（兜底，防止 SPA 外跳转）
        const origOpen = pageWindow.open;
        pageWindow.open = function (url) {
            if (url && typeof url === 'string' && url.includes('/istp-learning-center/')) {
                Storage.set('lastIndexUrl', url);
                location.href = url;
                return null;
            }
            return origOpen.apply(this, arguments);
        };
        try {
            cardEl.click();
        } finally {
            setTimeout(() => {
                pageWindow.open = origOpen;
            }, 2500);
        }
    }

    async function handleWorkHomePage() {
        decorateWorkHomeCards();
        if (!state.config.running) return;

        const cards = Array.from(document.querySelectorAll('.learningCourses .cardToLearn'));
        if (cards.length === 0) {
            log('正在等待「我学的课」列表加载...', 'info');
            return;
        }

        // 查找尚未完成（或尚未达标）的课程卡片
        const fullAuto = isFullAutoMastery();
        const threshold = clampChapterThreshold(state.config.fullAutoChapterThreshold);
        for (const card of cards) {
            const name = (card.querySelector('.courseName')?.innerText || '').trim();
            const pctText = (card.querySelector('.SelfProgress .percent')?.innerText || '').trim();
            const pctMatch = pctText.match(/(\d+)/);
            const pct = pctMatch ? parseInt(pctMatch[1], 10) : null;

            if (fullAuto) {
                const phaseA = state.config.fullAutoPhase !== 'mastery';

                // —— 阶段 A：刷章节，直到全部课程章节进度 ≥ 阈值 ——
                if (phaseA) {
                    if (name && state.chapterDoneCourses.has(name)) continue;
                    if (pct !== null && pct >= threshold) {
                        if (name) { state.chapterDoneCourses.add(name); saveChapterDoneCourses(); }
                        continue;
                    }
                    if (name && state.chapterRetriedCourses.has(name)) {
                        log(`【一键达标】课程【${name}】章节进度 ${pct === null ? '未知' : pct + '%'} 仍低于阈值 ${threshold}%，补扫后仍未达标，标记章节完成并继续`, 'warn');
                        state.chapterDoneCourses.add(name);
                        saveChapterDoneCourses();
                        continue;
                    }
                    if (name && state.completedCourses.has(name)) {
                        log(`【一键达标】课程【${name}】章节进度 ${pct === null ? '未知' : pct + '%'} 未达阈值 ${threshold}%，检测被略过的小节并补扫重试（已清空跳过缓存）`, 'info');
                        state.skipSections.clear();
                        saveSkipSections();
                        state.sectionFailCount.clear();
                        saveSectionFailCounts();
                        state.handledSections.clear();
                        saveHandledSections();
                        state.doneExercises.clear();
                        saveDoneExercises();
                        state.chapterRetriedCourses.add(name);
                        saveChapterRetriedCourses();
                    }
                    log(`【一键达标】阶段 A 进入课程：【${name || '课程'}】（章节进度 ${pct === null ? '未知' : pct + '%'}，阈值 ${threshold}%）`, 'info');
                    enterCourseCardInSameTab(card);
                    await sleep(3000);
                    return;
                }

                // —— 阶段 B：逐门循环刷掌握度到预设值 ——
                if (name && state.masteryMetCourses.has(name)) continue;
                log(`【一键达标】阶段 B 进入课程：【${name || '课程'}】（开始循环测评直到掌握度达标）`, 'info');
                enterCourseCardInSameTab(card);
                await sleep(3000);
                return;
            }

            if (name && state.completedCourses.has(name)) continue;
            if (pct !== null && pct >= 100) {
                state.completedCourses.add(name);
                saveCompletedCourses();
                continue;
            }

            if (state.config.autoNextCourse) {
                log(`自动进入待完成课程：【${name || '课程'}】`, 'info');
                enterCourseCardInSameTab(card);
                await sleep(3000);
                return;
            }
        }

        if (fullAuto) {
            // 本页课程已全部处理：先展开更多；展开完再按阶段推进（A→B 或结束）
            const moreBtn = document.querySelector('.learningCourses .expand .expand-btn');
            if (moreBtn && moreBtn.offsetWidth > 0 && (moreBtn.innerText || '').includes('展开')) {
                log('【一键达标】正在展开更多课程列表...', 'info');
                moreBtn.click();
                await sleep(2500);
                return;
            }
            if (state.config.fullAutoPhase !== 'mastery') {
                state.config.fullAutoPhase = 'mastery';
                saveConfig();
                updatePanelUI();
                log(`【一键达标】阶段 A（章节学习）全部课程已达 ${threshold}%，进入阶段 B：逐门循环刷掌握度到目标值`, 'success');
                return; // 下一轮主循环按新阶段重新选课
            }
            stopFullAutoMastery('🎉 【一键达标】全部课程均已达到目标掌握度，任务结束');
            return;
        }

        // 若当前页课程全部完成且存在「展开」更多按钮，自动展开下一页课程
        const expandBtn = document.querySelector('.learningCourses .expand .expand-btn');
        if (state.config.autoNextCourse && expandBtn && expandBtn.offsetWidth > 0 && (expandBtn.innerText || '').includes('展开')) {
            log('正在展开更多课程列表...', 'info');
            expandBtn.click();
            await sleep(2000);
            return;
        }

        log('请在课程卡片上点击「🚀 自动刷此课」，或开启「多课程连刷」自动进入课程。', 'info');
    }

    // =========================================================================
    // 6. 页面 B: 课程章节目录页 (/istp-learning-center/index)
    // =========================================================================
    function isExamTitle(title) {
        return /单元检测|单元测验|阶段检测|阶段测验|期中检测|期末检测|期末考试/i.test(title || '');
    }

    // 读取目录页当前课程名称：新版页面没有 .course-name-warp .name，
    // 取不到会导致 masteryMetCourses 写不进去 → 阶段 B 反复进入同一门课（实测死循环）
    function readCourseTitleOnIndexPage() {
        const sels = ['.course-name-warp .name', '.course-name', '.course-name-text',
            '[class*="course-name"]', '[class*="courseName"]'];
        for (const s of sels) {
            const el = document.querySelector(s);
            const t = (el && el.innerText ? el.innerText : '').trim().split('\n')[0].trim();
            if (t && t.length <= 40) return t;
        }
        return '';
    }

    async function handleCourseIndexPage() {
        // 记录当前课程目录页 URL，便于从章节页精准返回
        if (location.href.includes('/istp-learning-center/index')) {
            state.lastIndexUrl = location.href;
            Storage.set('lastIndexUrl', location.href);
        }

        if (!state.config.running) return;

        // 扫描期间钉住目录页：平台会不定时自动跳回"上次学习的小节"，打断目录树渲染
        state.indexPinUntil = Date.now() + 20000;

        state.currentChapterKey = '';
        state.processedInVisit.clear();

        // 1. 展开所有折叠的章节树节点
        const expandIcons = Array.from(document.querySelectorAll(
            '.index-tree-node .left-arrow:not(.is-open), .el-tree-node__expand-icon:not(.is-leaf):not(.expanded)'
        ));
        if (expandIcons.length > 0) {
            expandIcons.forEach(ic => ic.click());
            await sleep(1000);
        }

        // 2. 扫描所有小节行 (.aside-row)
        const rows = Array.from(document.querySelectorAll('.aside-row'));
        if (rows.length === 0) {
            log('正在等待课程章节树加载...', 'info');
            return;
        }

        // 补扫模式：彻底绕过所有跳过逻辑（hasCheck / skipSections / sessionSkipSections / handledSections），
        // 从头刷所有 leaf 小节，防止平台 Vue Router 恢复最后访问章节 + 本地过滤互相打架
        const forceRescan = !!state.forceRescanIndex;
        if (forceRescan) {
            log('【补扫模式】从头刷所有小节（跳过所有本地跳过/完成标记）', 'info');
        }

        const progEl = document.querySelector('.study-progress-top-left') ||
                       Array.from(document.querySelectorAll('*')).find(e => e.innerText && e.innerText.includes('学习总进度'));
        const totalProgress = progEl ? progEl.innerText.trim().replace(/\s+/g, ' ') : '';

        const todoList = [];
        let unrenderedRows = 0;
        rows.forEach((row, idx) => {
            const nameEl = row.querySelector('.tree-node-name');
            const resEl = row.querySelector('.right-content');
            if (!nameEl) return;

            const isLeaf = nameEl.classList.contains('no-child') || !row.querySelector('.left-arrow');
            if (!isLeaf) return; // 章节节点不参与待学判断

            // 目录树的状态列（完成图标 / 进度环 / 资源数）是异步填充的：行节点先出现，状态后补。
            // 若在填充前扫描，已完成的节会被当成待学 → 进去发现全做完 → 回目录 → 再误判，
            // 实测会反复进出同一批小节（每轮白跑约 6 秒）。这类行记为「未渲染」，本次不参与判断。
            if (!resEl || !resEl.querySelector('.icon-isfinished, .study-progress-wrapper, .progress-container, .number-of-resources')) {
                unrenderedRows++;
                return;
            }
            const resText = resEl.innerText.trim();
            const m = resText.match(/(\d+)\s*个(资源|学习)活动/);
            const count = m ? parseInt(m[1], 10) : 0;

            const hasCheck = !!resEl.querySelector('.icon-isfinished') ||
                hasStatusClassToken(resEl, /^(icon-)?is-?finish(ed)?$/i) ||
                /已完成|已学完|100%/.test(resText);
            const isLocked = !!row.querySelector('.icon-locked, .icon-becomedue') ||
                hasStatusClassToken(row, /^(icon-)?(is-)?locked$|^(icon-)?becomedue$/i);
            const name = nameEl.innerText.split('\n')[0].trim();
            const isExam = isExamTitle(name);

            if (isLeaf && count > 0) {
                // 锁定的始终跳过（平台不让学）
                if (isLocked) return;
                // 用户明确设置跳过单元检测 → 也跳过
                if (state.config.skipExams && isExam) { logSkipOnce(name, '单元检测（已开启跳过）'); return; }

                if (forceRescan) {
                    // 补扫模式：hasCheck / skipSections / sessionSkipSections / handledSections
                    // 全部绕过，从头刷
                    todoList.push({ idx, name, count, row });
                } else {
                    // 正常模式：保留所有过滤
                    if (hasCheck) return;
                    if (state.skipSections.has(name)) { logSkipOnce(name, '在永久跳过列表中'); return; }
                    if (state.sessionSkipSections.has(name)) { logSkipOnce(name, '本次运行已失败跳过，重新加载页面会重试'); return; }
                    if (state.handledSections.has(sectionDoneKey(name))) { logSkipOnce(name, '剩余练习平台已判定掌握，本地标记为已处理'); return; }
                    todoList.push({ idx, name, count, row });
                }
            }
        });

        // 补扫模式只对这一次扫描生效，用完即清
        state.forceRescanIndex = false;

        if (todoList.length > 0) {
            // 扫到了待学小节，重置空扫计数
            state.indexEmptyScanCount = 0;
            state.indexRenderWait = 0;
        } else if (unrenderedRows > 0) {
            // 还有叶子小节的状态列没渲染出来：先等它填充，避免把「已完成」误判成「待学」
            const waited = (state.indexRenderWait || 0) + 1;
            state.indexRenderWait = waited;
            if (waited <= 6) {
                log(`目录状态列尚未渲染完成（${unrenderedRows} 行待填充），稍后重扫（第 ${waited} 次）`, 'info');
                await sleep(1200);
                return;
            }
            state.indexRenderWait = 0;
        }

        if (todoList.length === 0) {
            // 防御：章节树是异步渲染的，刚切回目录页时子节点可能还没展开/插入 DOM，
            // 此时会误判「全部刷完」→ 回首页 → 重新进课，白跑一整轮。
            // 因此连续空扫 3 次（每次再点一次展开 + 等 2 秒）才认定真的刷完。
            const emptyRetried = state.indexEmptyScanCount || 0;
            if (!forceRescan && emptyRetried < 3) {
                state.indexEmptyScanCount = emptyRetried + 1;
                log(`目录暂未发现待完成小节，可能仍在渲染，重试扫描（第 ${emptyRetried + 1} 次）`, 'warn');
                const retryIcons = Array.from(document.querySelectorAll(
                    '.index-tree-node .left-arrow:not(.is-open), .el-tree-node__expand-icon:not(.is-leaf):not(.expanded)'
                ));
                if (retryIcons.length > 0) retryIcons.forEach(ic => ic.click());
                await sleep(2000);
                return;
            }
            state.indexEmptyScanCount = 0;

            const courseTitle = readCourseTitleOnIndexPage();
            log(`🎉 当前课程可学小节已全部刷完！（${totalProgress || '总进度未知，请回首页核对'}）`, 'success');

            // 私有「一键达标」两阶段：阶段 A 只回首页判进度；阶段 B 循环刷掌握度到预设值
            if (isFullAutoMastery() && state.config.running) {
                if (state.config.fullAutoPhase !== 'mastery') {
                    // 阶段 A：章节刷完即返回学习空间，由首页按进度判定是否达标 / 补扫
                    // 跳首页前清掉持久化的小节级/活动级本地标记，避免补扫时把中间节误判为"已处理"
                    // （平台可能把"已练过"但未真正学习的活动在小节页快速扫过，写入 handledSections，
                    // 导致下次进目录页时这些节被跳过 → todoList 永远为空 → 反复横跳）
                    state.handledSections.clear();
                    saveHandledSections();
                    state.doneExercises.clear();
                    saveDoneExercises();
                    state.sessionFailedExercises.clear();
                    state.exerciseFailCount.clear();
                    state.platformDoneVisits.clear();
                    state.sessionSkipSections.clear();
                    state.currentChapterKey = '';
                    state.processedInVisit.clear();
                    // 允许首页重新判定这门课，不提前锁死"已完成"
                    if (courseTitle) {
                        state.completedCourses.delete(courseTitle);
                        saveCompletedCourses();
                        state.chapterDoneCourses.delete(courseTitle);
                        saveChapterDoneCourses();
                        state.chapterRetriedCourses.delete(courseTitle);
                        saveChapterRetriedCourses();
                    }
                    log('【一键达标】阶段 A：本课程可学小节已刷完，已清理本地跳过/完成标记，返回学习空间检查章节进度', 'info');
                    await sleep(1500);
                    location.href = 'https://icourse.fifedu.com/istp-manage/work/home';
                    return;
                }

                // 阶段 B：循环刷掌握度到预设值
                if (state.evaluationRunning) return;
                if (!courseTitle) {
                    // 极端情况：拿不到课程名就无法记录「已达标」，继续会无限重复阶段 B，这里直接收尾
                    log('未能读取课程名称，无法记录掌握度达标状态，为避免死循环已结束一键达标', 'warn');
                    state.config.fullAutoMastery = false;
                    state.config.running = false;
                    saveConfig();
                    updatePanelUI();
                    return;
                }
                if (courseTitle && state.masteryMetCourses.has(courseTitle)) {
                    // 已达标，直接进入下一门，避免死循环
                } else {
                    const parsed = parseEvalTargetInput(state.config.evalTargetMastery);
                    // 轮次上限：默认 8 轮；用户开启「无限循环」或设置了合法目标值则尊重其配置
                    let phaseBMaxRounds = 8;
                    if (state.config.evalInfinite) phaseBMaxRounds = 500;
                    else if (parsed.provided && parsed.valid) phaseBMaxRounds = 20;
                    if (!parsed.provided || !parsed.valid) {
                        log(`【一键达标】阶段 B：未设置合法目标掌握度，最多测评 ${phaseBMaxRounds} 轮后切换下一门`, 'warn');
                    } else {
                        log(`【一键达标】阶段 B：开始循环测评，目标掌握度 ${parsed.value}%（最多 ${phaseBMaxRounds} 轮）`, 'info');
                    }
                    await runCourseEvaluation({ maxRounds: phaseBMaxRounds });
                    if (courseTitle) {
                        state.masteryMetCourses.add(courseTitle);
                        saveMasteryMetCourses();
                    }
                    if (!state.config.running || state.evalStopRequested) {
                        log('【一键达标】测评被手动停止，任务结束', 'warn');
                        state.config.fullAutoMastery = false;
                        saveConfig();
                        updatePanelUI();
                        return;
                    }
                    log('【一键达标】本课程阶段 B 测评流程已结束，等待答题界面关闭后返回「学习空间」', 'success');
                }
                await waitEvaluationUiClosed();
                await sleep(2000);
                location.href = 'https://icourse.fifedu.com/istp-manage/work/home';
                return;
            }

            if (state.config.autoNextCourse) {
                log('3 秒后自动返回「学习空间」切换下一门未完成课程...', 'info');
                await sleep(3000);
                location.href = 'https://icourse.fifedu.com/istp-manage/work/home';
            } else {
                state.config.running = false;
                saveConfig();
                updatePanelUI();
            }
            return;
        }

        const target = todoList[0];
        log(`准备进入小节：【${target.name}】(剩余 ${todoList.length} 节未完成)`, 'info');

        // 存目标小节信息，让学习页校验是否真进到了这一节
        state.targetSectionName = target.name;
        state.targetSectionChapterId = '';
        try {
            const vueComp = target.row.__vue__ || target.row.querySelector('.tree-node-name')?.__vue__;
            const data = vueComp?.node?.data || vueComp?.row?.data || vueComp?.data;
            state.targetSectionChapterId = String(
                data?.chapterId || data?.id || target.row.dataset?.chapterId || target.row.dataset?.id || ''
            );
        } catch (e) { /* ignore */ }

        // 记录点之前的 URL，等 Vue Router push 生效
        const prevPath = location.pathname;
        const clickTarget = target.row.querySelector('.tree-node-name') || target.row;
        // 主动进入小节：解除目录页钉住，否则 Router 守卫会把本次跳转反弹回 /index
        state.indexPinUntil = 0;
        fireHumanClick(clickTarget);

        // 必须等到 URL 真的变了（Vue Router push 完成）才 return，
        // 否则下一轮 mainTick 又进 handleCourseIndexPage 又点一次 → 多次 push 混乱
        const waitStart = Date.now();
        let pushed = false;
        while (Date.now() - waitStart < 8000 && state.config.running) {
            if (location.pathname.includes('/pattern/chapter')) {
                pushed = true;
                break;
            }
            if (location.pathname !== prevPath) {
                pushed = true;  // 任何变化都算（可能跳了别的路径）
                break;
            }
            await sleep(400);
        }
        if (!pushed) {
            log('点击小节后 8 秒内 Vue Router 未跳转，再点一次', 'warn');
            fireHumanClick(clickTarget);
            await sleep(1500);
        }
    }

    // =========================================================================
    // 7. 页面 C: 章节学习详情页 (/istp-learning-center/pattern/chapter)
    // =========================================================================
    async function returnToCourseIndex() {
        state.processedInVisit.clear();
        state.currentChapterKey = '';
        state.indexEmptyScanCount = 0;
        state.indexRenderWait = 0;
        // 标记为「主动离开」，让 confirm 拦截器放行导航，而不是留在练习页
        state.leavingIntentionally = true;
        setTimeout(() => { state.leavingIntentionally = false; }, 5000);

        // 优先点击页面左上角返回按钮 (保留 Vue Router 状态)
        const backBtn = document.querySelector('.study-header .return-icon, .page-return-p .return-icon, .back-btn');
        if (backBtn) {
            backBtn.click();
            await sleep(2500);
            if (location.pathname.includes('/istp-learning-center/index')) return;
        }

        // 其次使用 Vue Router 导航回 /index
        const router = getVueRouter();
        const params = new URLSearchParams(location.search);
        const courseId = params.get('courseId') || (sessionStorage.getItem('courseId') || '').replace(/^"|"$/g, '');
        const classId = params.get('classId') || (sessionStorage.getItem('classId') || '').replace(/^"|"$/g, '');
        const termId = params.get('termId') || (sessionStorage.getItem('termId') || '').replace(/^"|"$/g, '');

        if (router && courseId) {
            router.push({
                path: '/index',
                query: { courseId, classId, termId }
            });
            // 等 Vue Router push 生效，最多 6 秒
            const start = Date.now();
            while (Date.now() - start < 6000) {
                if (location.pathname.includes('/istp-learning-center/index')) return;
                await sleep(300);
            }
            return;
        }

        if (state.lastIndexUrl) {
            location.href = state.lastIndexUrl;
        } else {
            location.href = `https://icourse.fifedu.com/istp-learning-center/index?courseId=${courseId}&classId=${classId}&termId=${termId}`;
        }
        await sleep(3000);
    }

    function inspectCurrentChapterDOM() {
        const curChapter = (document.querySelector('.current-chapter-name, .chapter-name-title')?.innerText || '')
            .split('\n')[0].trim();

        const itemEls = Array.from(document.querySelectorAll('.activity-list-item'));
        const items = itemEls.map((el, i) => {
            const nameEl = el.querySelector('.activity-name');
            const name = (nameEl?.innerText || '').replace(/选学/g, '').trim();
            const statusText = (el.querySelector('.status-flex')?.innerText || el.innerText || '')
                .replace(/\s+/g, ' ').trim();
            const isActive = el.classList.contains('is-active') || !!el.querySelector('.is-current');
            const statusEl = el.querySelector('.status-flex');
            const doneKey = exerciseKey(curChapter, name);
            const isDone = state.doneExercises.has(doneKey);
            const isSessionFailed = state.sessionFailedExercises.has(doneKey);
            // 仅来自平台页面自身的完成标记（不包含本地记录），用于判断「平台侧确实都做完了」
            const isPlatformDone = !!el.querySelector('.is-finish, .is-finished') ||
                hasStatusClassToken(el, /^(icon-)?is-?finish(ed)?$/i) ||
                statusText.includes('已完成') ||
                (!!statusEl && /100%/.test(statusEl.innerText || ''));
            const isFinish = isDone || isSessionFailed || isPlatformDone;
            return { i, name, statusText, isActive, isFinish, isDone, isSessionFailed, isPlatformDone, el };
        });

        const activeIdx = items.findIndex(x => x.isActive);
        const curItem = activeIdx >= 0 ? items[activeIdx] : items[0] || null;
        const curName = curItem ? curItem.name : (document.querySelector('.chapter-activity-title, h3')?.innerText || '').trim();
        const statusStr = curItem ? curItem.statusText : '';

        const isExam = isExamTitle(curChapter) || isExamTitle(curName);
        const video = document.querySelector('video, audio');

        const isPPT = /\.(ppt|pptx)($|\s)/i.test(curName) ||
                      statusStr.includes('PPT') ||
                      !!document.querySelector('.pptWrap, .pptContainer, .officeweb365-iframe, .pptIframe');

        const isDoc = !isPPT && (
            /\.(doc|docx|pdf|txt|xls|xlsx)($|\s)/i.test(curName) ||
            /DOC|PDF|Word|文档|表格/i.test(statusStr) ||
            !!document.querySelector('.pdfContainer, .docWrap, .wordWrap, .pdfWrap, iframe[src*="viewer.html"]')
        );

        const isVideo = !isDoc && !isPPT && (!!video || /视频|音频|\.(mp4|webm|mp3|wav)/i.test(curName + ' ' + statusStr));

        const isExercise = !isExam && !isDoc && !isPPT && !isVideo && (
            !!document.querySelector('.SubjectLib, .SubjectCover, .SubjectLayout, .CloseBtn') ||
            /练习|自适应/i.test(statusStr) ||
            Array.from(document.querySelectorAll('button, .el-button')).some(b => {
                const t = (b.innerText || '').trim();
                return t.includes('开始练习') || t.includes('继续练习') || t.includes('选完了');
            })
        );

        const isOtherInstant = !isExam && !isDoc && !isPPT && !isVideo && !isExercise &&
                               !!document.querySelector('.previewContainer, .homework_box');

        return {
            curChapter,
            curName,
            items,
            totalItems: items.length,
            activeIdx: activeIdx >= 0 ? activeIdx : 0,
            isExam,
            isVideo,
            isPPT,
            isDoc,
            isExercise,
            isOtherInstant
        };
    }

    /**
     * 处理视频/音频课件
     */
    async function processVideoActivity(curName) {
        log(`正在处理音视频课件：【${curName}】`, 'info');
        const media = document.querySelector('video, audio');

        if (media) {
            try {
                media.muted = true;
                media.playbackRate = Math.min(Math.max(state.config.videoRate || 4.0, 1.0), 16.0);
                media.play().catch(() => {});
            } catch (e) {}
            const playBtn = document.querySelector('.xgplayer-start, .xgplayer-play, .xg-icon-play');
            if (playBtn) playBtn.click();
        }

        // 等待元数据加载 (最多 8 秒)，同时检测格式错误
        let metadataLoaded = false;
        let hasMediaError = false;
        for (let i = 0; i < 16; i++) {
            if (!state.config.running) return;
            const v = document.querySelector('video, audio');
            const errEl = document.querySelector('.xgplayer-error, .xgplayer-error-text');
            const errVisible = errEl && errEl.offsetWidth > 0 && (errEl.innerText || '').includes('不支持');

            if ((v && v.error) || errVisible) {
                hasMediaError = true;
                break;
            }
            if (v && Number.isFinite(v.duration) && v.duration > 0) {
                metadataLoaded = true;
                break;
            }
            await sleep(500);
        }

        // 情况 1: 视频源损坏或开启了极速模式 -> 直接通过 Vue 进度接口上报 100%
        if (hasMediaError || !metadataLoaded || state.config.fastMode) {
            const v = document.querySelector('video, audio');
            const totalDur = (v && Number.isFinite(v.duration) && v.duration > 0) ? Math.floor(v.duration) : 300;
            if (hasMediaError) {
                log(`检测到平台视频源异常，启用底层进度直报修复：【${curName}】`, 'warn');
            } else {
                log(`正在同步视频完成进度 (时长: ${totalDur}s)...`, 'info');
            }
            triggerVueStudyProcess({ total: totalDur, end: totalDur, type: '275002' });
            if (v && Number.isFinite(v.duration) && v.duration > 2) {
                try {
                    v.currentTime = Math.max(0, v.duration - 1);
                    v.dispatchEvent(new Event('seeked'));
                    v.dispatchEvent(new Event('ended'));
                } catch (e) {}
            }
            await sleep(3000);
            return;
        }

        // 情况 2: 常规倍速 + 步进快进模式
        let steps = 0;
        while (steps < 180 && state.config.running) {
            const v = document.querySelector('video, audio');
            if (!v) break;
            const dur = Number.isFinite(v.duration) ? v.duration : 0;
            const cur = Number.isFinite(v.currentTime) ? v.currentTime : 0;

            if (v.ended || (dur > 0 && cur >= dur - 2)) {
                log(`视频播放完成 (${Math.round(cur)}/${Math.round(dur)}s)`, 'success');
                break;
            }
            if (v.paused) {
                v.muted = true;
                v.play().catch(() => {});
            }
            v.playbackRate = Math.min(Math.max(state.config.videoRate || 4.0, 1.0), 16.0);
            try {
                v.currentTime = Math.min(dur - 0.5, cur + 15);
            } catch (e) {}

            const player = document.querySelector('.xgplayer') || v;
            player.dispatchEvent(new KeyboardEvent('keydown', {
                key: 'ArrowRight', code: 'ArrowRight', keyCode: 39, bubbles: true
            }));

            if (steps % 6 === 0 && dur > 0) {
                log(`视频进度：${Math.round((cur / dur) * 100)}% (${Math.round(cur)}/${Math.round(dur)}s)`);
            }
            steps++;
            await sleep(400);
        }

        const v = document.querySelector('video, audio');
        const finalDur = (v && Number.isFinite(v.duration) && v.duration > 0) ? Math.floor(v.duration) : 300;
        triggerVueStudyProcess({ total: finalDur, end: finalDur, type: '275002' });
        await sleep(2500);
    }

    /**
     * 读取 PPT 当前页/总页数：优先取子框架 postMessage，其次取页面页码文本
     */
    function readPptPageState() {
        let current = 0, total = 0;
        if (pptFrameState && pptFrameState.at && Date.now() - pptFrameState.at < 5000) {
            current = pptFrameState.current || 0;
            total = pptFrameState.total || 0;
        }
        if (!current || !total) {
            const scope = document.querySelector('.pptContainer, .pptWrap, .previewContainer, .btnWrap') || document;
            const m = (scope.innerText || '').match(/(\d{1,3})\s*[\/／]\s*(\d{1,3})/);
            if (m) {
                const c = parseInt(m[1], 10), t = parseInt(m[2], 10);
                if (t > 0 && c > 0 && c <= t) { current = current || c; total = total || t; }
            }
        }
        return { current, total };
    }

    /**
     * 尝试推进 PPT 一页：按钮点击 + 键盘右箭头 + postMessage 三种方式兜底
     */
    function advancePptPage(nextPage, totalPages) {
        const nextBtn = document.querySelector('.pptContainer .toggle .pptBtn:last-child, .btnWrap .arrow-right, .btnWrap .pptBtn:last-child, [title*="下一页"], [title*="下一张"]');
        if (nextBtn) {
            try { nextBtn.click(); } catch (e) {}
        }
        const scope = document.querySelector('.pptContainer, .pptWrap, .previewContainer') || document.body;
        try {
            scope.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', code: 'ArrowRight', keyCode: 39, bubbles: true }));
        } catch (e) {}
        // 上报给平台时总页数必须使用真实总页数，且打上自身标记，避免被自己的监听器读走
        const total = totalPages > 0 ? totalPages : nextPage;
        pageWindow.postMessage({
            currentPageNum: nextPage,
            totalPageNum: total,
            data: { currentPageNum: nextPage, totalPageNum: total },
            __fifSelf: pageWindow.__fifHelperInstanceId
        }, '*');
    }

    /**
     * 处理 PPT 课件：按真实页码逐页翻到末页后再上报完成，平台卡顿时也不会中途放弃
     */
    async function processPPTActivity(curName) {
        log(`正在处理 PPT 课件：【${curName}】`, 'info');

        // 每个课件重新开始统计页码，避免上一个课件的残留页码串到本课件
        pptFrameState.current = 0;
        pptFrameState.total = 0;
        pptFrameState.at = 0;

        // —— 阶段 0：等 PPT 子框架 / 容器真正加载出来 ——
        let { current, total: totalPages } = { current: 0, total: 0 };
        let pptReady = false;
        for (let i = 0; i < 30; i++) {
            const iframe = document.querySelector('.officeweb365-iframe, iframe[src*="ow365"], .pptIframe');
            const container = document.querySelector('.pptContainer, .pptWrap, .previewContainer');
            if ((iframe && iframe.offsetWidth > 0) || (container && container.offsetWidth > 0)) {
                const st = readPptPageState();
                if (st.current > 0 && st.total > 0) {
                    current = st.current;
                    totalPages = st.total;
                    pptReady = true;
                    break;
                }
            }
            if (i === 3 || i === 9) log(`PPT 等待加载中（${(i + 1) * 500}ms）...`, 'info');
            await sleep(500);
        }

        if (!pptReady) {
            log('PPT 加载超时，无法获取真实页码，暂不上报完成（避免误判）', 'warn');
            await sleep(2000);
            return;
        }

        log(`PPT 加载就绪，初始页码：${current}/${totalPages}`, 'info');

        // retry 策略：最多 3 轮完整翻页循环，每轮 stuck 后等 3-5 秒让平台恢复
        const MAX_RETRIES = 3;
        let retry = 0;
        let lastKnownPage = current;  // 必须在上报前记住真实进度，不能用上报后的状态
        let reachedEnd = current >= totalPages;

        while (retry < MAX_RETRIES && !reachedEnd) {
            const deadline = Date.now() + Math.min(180000, Math.max(30000, totalPages * 4000));
            let stuck = 0, iters = 0;
            let loopReachedEnd = false;

            while (state.config.running && Date.now() < deadline && !loopReachedEnd) {
                iters++;
                const before = readPptPageState();
                if (before.total > 0) totalPages = Math.max(totalPages, before.total);
                if (before.current > 0) current = before.current;
                lastKnownPage = current;  // 每次循环开头更新真实进度

                if (totalPages > 0 && current >= totalPages) { loopReachedEnd = true; break; }

                advancePptPage(current + 1, totalPages);
                // 等页码真的推进：最多 4 秒，每 300ms 查一次
                let curAfter = current, totalAfter = totalPages, settled = false;
                for (let w = 0; w < 14; w++) {
                    await sleep(300);
                    const st = readPptPageState();
                    if (st.total > 0) totalAfter = Math.max(totalAfter, st.total);
                    if (st.current > current) { curAfter = st.current; settled = true; break; }
                }

                if (settled && curAfter > current) {
                    current = curAfter;
                    totalPages = totalAfter > 0 ? totalAfter : totalPages;
                    stuck = 0;
                    if (iters % 3 === 0) log(`PPT 翻页中：${current}/${totalPages}`, 'info');
                } else {
                    stuck++;
                }
                if (totalPages > 0 && curAfter >= totalPages) { loopReachedEnd = true; break; }
                if (stuck >= 8) {
                    // 这一轮翻页卡住了，不直接报完成 → 记下真实进度，retry 等一会再试
                    log(`PPT 翻页卡住（${stuck} 次未推进），当前真实进度 ${current}/${totalPages}，准备第 ${retry + 1} 轮重试...`, 'warn');
                    break;
                }
            }

            if (loopReachedEnd) {
                reachedEnd = true;
                lastKnownPage = totalPages;
                break;
            }

            retry++;
            if (retry < MAX_RETRIES) {
                // retry 前先上报当前已真实到达的进度（不是末页！）
                log(`上报当前真实进度 ${lastKnownPage}/${totalPages}，等待 5 秒后重试翻页...`, 'info');
                for (let i = 0; i < 3; i++) {
                    triggerVueStudyProcess({ total: totalPages, end: lastKnownPage, type: '275003' });
                    await sleep(400);
                }
                // 清空 postMessage 上报的旧状态，让平台重新初始化 PPT 组件
                try {
                    const container = document.querySelector('.pptContainer, .pptWrap, .previewContainer');
                    if (container) {
                        // 强制触发重渲染：先 detach 再 appendChild
                        const parent = container.parentNode;
                        if (parent) {
                            parent.removeChild(container);
                            await sleep(300);
                            parent.appendChild(container);
                            await sleep(800);
                        }
                    }
                } catch (e) { /* ignore */ }
                await sleep(5000);
                // 重新读一次初始状态
                const st = readPptPageState();
                current = st.current > 0 ? st.current : lastKnownPage;
                if (st.total > 0) totalPages = st.total;
            }
        }

        // —— 最终上报 ——
        // reachedEnd=true → 真实翻到末页，上报 100%
        // 否则 → 上报最后一次真实进度（不再硬拉到末页），不标记"完成"
        // 单页课件没有翻页动作，平台需要足够的停留时间才记录完成，这里补足停留
        if (reachedEnd && totalPages <= 1) await sleep(8000);
        // 极速模式下，个别课件（内嵌视频/动画的幻灯片）无法程序翻页，
        // 为避免整节卡在 <100%，强制上报完成交给平台判定；关闭极速模式则保持真实进度
        const forceComplete = !reachedEnd && state.config.fastMode && totalPages > 0;
        const reportEnd = (reachedEnd || forceComplete) ? totalPages : lastKnownPage;
        const pct = totalPages > 0 ? Math.round((lastKnownPage / totalPages) * 100) : 0;
        for (let i = 0; i < 3; i++) {
            triggerVueStudyProcess({ total: totalPages, end: reportEnd, type: '275003' });
            await sleep(400);
        }

        if (reachedEnd) {
            log(`PPT 课件已完成：【${curName}】（${lastKnownPage}/${totalPages}，约 ${pct}%）`, 'success');
        } else if (forceComplete) {
            log(`PPT 翻页受阻（真实进度 ${lastKnownPage}/${totalPages}），极速模式已强制上报完成：【${curName}】`, 'warn');
        } else {
            log(`PPT 课件未能翻到末页：【${curName}】（真实进度 ${lastKnownPage}/${totalPages}，约 ${pct}%，已上报真实进度，等待平台自行判定）`, 'warn');
        }
        await sleep(1500);
    }

    /**
     * 处理 Word / PDF / TXT / Excel 文档课件
     */
    async function processDocActivity(curName) {
        log(`正在处理文档/PDF课件：【${curName}】`, 'info');
        await sleep(1500);

        // 1. 向 PDF.js iframe 发送跳转末页消息
        const pdfIframe = document.querySelector('.pdfContainer iframe, iframe[src*="viewer.html"]');
        if (pdfIframe && pdfIframe.contentWindow) {
            for (const p of [5, 15, 30, 60, 100]) {
                pdfIframe.contentWindow.postMessage({ pageNum: p }, '*');
            }
        }

        // 2. 滚动页面内部所有可滚动容器到底部
        const scrollables = Array.from(document.querySelectorAll('.docWrap, .wordWrap, .pdfWrap, .wrapBox, .main-warp .el-scrollbar__wrap'));
        scrollables.forEach(el => {
            if (el.scrollHeight > el.clientHeight) {
                el.scrollTop = el.scrollHeight;
                el.dispatchEvent(new Event('scroll'));
            }
        });

        // 3. 直接触发 Vue 组件 100% 阅读完成上报 (彻底解决单页 PDF 不打勾 Bug)
        triggerVueStudyProcess({ total: 10, end: 10, type: '275003' });
        log(`文档已置底并同步完成状态：【${curName}】`, 'success');
        await sleep(3500);
    }

    // =========================================================================
    // 6.5 AI 自动答题模块（DeepSeek / OpenAI 兼容接口）
    // =========================================================================
    const AI_TAB_ID = Math.random().toString(36).slice(2, 10);
    const AI_LOCK_KEY = 'aiAnswerLock';
    const AI_CACHE_KEY = 'aiAnswerCache';
    let aiLastRequestAt = 0;
    const aiAnswerCache = Storage.get(AI_CACHE_KEY, {}) || {};

    function clampNumber(value, min, max, fallback) {
        const n = Number(value);
        if (!Number.isFinite(n)) return fallback;
        return Math.min(max, Math.max(min, n));
    }

    function hashString(str) {
        let h = 5381;
        for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0;
        return (h >>> 0).toString(36);
    }

    function htmlToPlainText(html) {
        if (!html) return '';
        const div = document.createElement('div');
        div.innerHTML = String(html);
        div.querySelectorAll('script, style').forEach(el => el.remove());
        return (div.textContent || '')
            .replace(/\u00a0/g, ' ')
            .replace(/[ \t]+\n/g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
    }

    function isVisibleEl(el) {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        if (rect.width <= 0 || rect.height <= 0) return false;
        const style = getComputedStyle(el);
        return style.display !== 'none' && style.visibility !== 'hidden';
    }

    // 在文档、同源 iframe、Shadow DOM 中查找练习根组件（优先 .SubjectLib，其次按 Vue 实例特征匹配）
    function collectDocuments(root = document, out = [], depth = 0) {
        if (!root || depth > 3) return out;
        out.push(root);
        let frames = [];
        try { frames = Array.from(root.querySelectorAll('iframe')); } catch (e) { frames = []; }
        for (const f of frames) {
            try {
                const doc = f.contentDocument;
                if (doc) collectDocuments(doc, out, depth + 1);
            } catch (e) { /* 跨域忽略 */ }
        }
        return out;
    }

    const PRACTICE_SEED_SELECTORS = ['.SubjectLib', '.subjectcon', '.SubjectCon', '.dialogBox', '.subject', '.question', '.RCJTmp', '.flatLay', '.Cloze'];

    function collectPracticeSeeds(root) {
        const seeds = [];
        try {
            for (const sel of PRACTICE_SEED_SELECTORS) root.querySelectorAll(sel).forEach(el => seeds.push(el));
            const all = root.querySelectorAll('*');
            for (let i = 0; i < all.length && i < 6000; i++) {
                const sr = all[i].shadowRoot;
                if (sr) for (const sel of PRACTICE_SEED_SELECTORS) sr.querySelectorAll(sel).forEach(el => seeds.push(el));
            }
        } catch (e) { /* ignore */ }
        return seeds;
    }

    function findInstanceFromSeeds(seeds) {
        for (const seed of seeds) {
            let cur = seed && seed.__vueParentComponent;
            let guard = 0;
            while (cur && guard++ < 40) {
                const st = cur.setupState;
                if (st && st.Question && typeof st.inspect === 'function') return cur;
                cur = cur.parent;
            }
        }
        return null;
    }

    function getSubjectLibInstance() {
        const docs = collectDocuments();
        for (const doc of docs) {
            try {
                const el = doc.querySelector('.SubjectLib');
                if (el) {
                    let inst = el.__vueParentComponent;
                    while (inst && !(inst.type && inst.type.name === 'SubjectLib')) inst = inst.parent;
                    if (inst && inst.setupState) return inst;
                }
            } catch (e) { /* ignore */ }
        }
        const seeds = [];
        for (const doc of docs) {
            try { seeds.push(...collectPracticeSeeds(doc.body || doc)); } catch (e) { /* ignore */ }
        }
        const hit = findInstanceFromSeeds(seeds);
        if (hit) return hit;
        // 兜底：遍历所有元素，查找带 Question/inspect 的 Vue 组件实例
        for (const doc of docs) {
            try {
                const all = Array.from(doc.querySelectorAll('*')).slice(0, 6000);
                const hit2 = findInstanceFromSeeds(all);
                if (hit2) return hit2;
            } catch (e) { /* ignore */ }
        }
        return null;
    }

    // 找不到练习界面时输出诊断信息，便于定位实际 DOM 结构
    function describePracticeDom() {
        try {
            const cnt = (sel) => { try { return document.querySelectorAll(sel).length; } catch (e) { return -1; } };
            const parts = [];
            parts.push('iframe=' + cnt('iframe'));
            parts.push('subjectcon=' + cnt('.subjectcon') + ',SubjectCon=' + cnt('.SubjectCon') + ',dialogBox=' + cnt('.dialogBox') + ',subject=' + cnt('.subject'));
            const cls = Array.from(new Set(Array.from(document.querySelectorAll('[class]'))
                .map(e => String(e.className || ''))
                .filter(c => /subject|Subject|question|Question|dialog|Dialog|cover|Cover|lib|Lib/.test(c))));
            if (cls.length) parts.push('候选类名=' + cls.slice(0, 10).join('|'));
            const lib = document.querySelector('.SubjectLib');
            if (lib) {
                const outline = Array.from(lib.querySelectorAll('*')).slice(0, 28).map(e => {
                    const cn = typeof e.className === 'string' ? e.className.trim().split(/\s+/).slice(0, 2).join('.') : '';
                    return e.tagName.toLowerCase() + (cn ? '.' + cn : '');
                }).join(',');
                parts.push('lib节点=' + outline);
                const btns = Array.from(lib.querySelectorAll('button')).map(b => String(b.innerText || '').trim()).filter(Boolean);
                if (btns.length) parts.push('按钮=' + btns.join('|'));
                const lt = String(lib.innerText || '').replace(/\s+/g, ' ').slice(0, 100);
                if (lt) parts.push('lib文本=' + lt);
            }
            const txt = String(document.body.innerText || '').replace(/\s+/g, ' ').slice(0, 60);
            if (txt) parts.push('页面文本=' + txt);
            return parts.join('；');
        } catch (e) {
            return '诊断失败';
        }
    }

    // 诊断通道：把练习界面快照上传到本机服务 (127.0.0.1:18765/diag)，便于离线分析 DOM
    function uploadDiag(payload) {
        try {
            const body = JSON.stringify(payload);
            if (typeof GM_xmlhttpRequest === 'function') {
                GM_xmlhttpRequest({
                    method: 'POST',
                    url: 'http://127.0.0.1:18765/diag',
                    headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
                    data: body,
                    timeout: 6000,
                    onload: () => {},
                    onerror: () => {}
                });
            } else {
                fetch('http://127.0.0.1:18765/diag', { method: 'POST', body }).catch(() => {});
            }
        } catch (e) { /* ignore */ }
    }

    function snapshotPracticeDOM(tag) {
        const snap = { tag, ts: new Date().toLocaleTimeString('zh-CN', { hour12: false }), url: String(location.href).slice(0, 200), docs: [] };
        try {
            for (const doc of collectDocuments()) {
                const d = { url: '', lib: false };
                try { d.url = String((doc.location && doc.location.href) || '').slice(0, 160); } catch (e) { d.url = '(cross-origin)'; }
                let lib = null;
                try { lib = doc.querySelector('.SubjectLib'); } catch (e) { /* ignore */ }
                d.lib = !!lib;
                const root = lib || (doc.body || null);
                if (root) {
                    try { d.text = String(root.innerText || '').replace(/\s+/g, ' ').slice(0, 600); } catch (e) { /* ignore */ }
                    try {
                        d.nodes = Array.from(root.querySelectorAll('*')).slice(0, 70).map(e => {
                            const cn = typeof e.className === 'string' ? e.className.trim().split(/\s+/).slice(0, 3).join('.') : '';
                            return e.tagName.toLowerCase() + (cn ? '.' + cn : '');
                        }).join(',');
                    } catch (e) { /* ignore */ }
                    try {
                        d.buttons = Array.from(root.querySelectorAll('button, .el-button, [class*="btn"]'))
                            .map(b => String(b.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 20))
                            .filter(Boolean).slice(0, 30);
                    } catch (e) { /* ignore */ }
                    try { d.html = String(root.outerHTML || '').replace(/\s+/g, ' ').slice(0, 5000); } catch (e) { /* ignore */ }
                }
                snap.docs.push(d);
            }
            try {
                snap.overlays = Array.from(document.querySelectorAll('.el-overlay, .el-dialog, .el-message-box, .CEDialogModal'))
                    .filter(isVisibleEl)
                    .map(o => String(o.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 200))
                    .filter(Boolean);
            } catch (e) { /* ignore */ }
        } catch (e) { /* ignore */ }
        uploadDiag(snap);
        return snap;
    }

    // 网络探针：临时 Hook fetch/XHR，把练习相关接口的请求与响应上报到本机诊断服务
    let netTapState = null;

    function netTapShouldLog(url) {
        return /subject|question|answer|practice|exercis|adaptive|knowledge|report|batch|activity|chapterApi|courseApi|topic/i.test(String(url));
    }

    function startNetTap(tag) {
        stopNetTap();
        const win = pageWindow;
        const tap = {
            tag,
            records: [],
            origFetch: win.fetch,
            origOpen: win.XMLHttpRequest.prototype.open,
            origSend: win.XMLHttpRequest.prototype.send,
            upload: function (rec) {
                tap.records.push(rec);
                if (tap.records.length > 80) tap.records.shift();
                try { inspectTapRecord(rec); } catch (e) { /* ignore */ }
                uploadDiag({ tag: 'net', netTag: tag, ts: new Date().toLocaleTimeString('zh-CN', { hour12: false }), rec });
            }
        };
        netTapState = tap;

        try {
            win.fetch = function (input, init) {
                const url = typeof input === 'string' ? input : (input && input.url) || '';
                const promise = tap.origFetch.apply(this, arguments);
                if (netTapShouldLog(url)) {
                    try {
                        promise.then(res => {
                            try {
                                res.clone().text().then(t => {
                                    tap.upload({
                                        via: 'fetch', url: String(url).slice(0, 400),
                                        method: (init && init.method) || (input && input.method) || 'GET',
                                        req: init && init.body ? String(init.body).slice(0, 1500) : '',
                                        status: res.status, res: String(t).slice(0, 20000)
                                    });
                                }).catch(() => {});
                            } catch (e) { /* ignore */ }
                        }).catch(err => {
                            tap.upload({ via: 'fetch', url: String(url).slice(0, 400), error: String(err && err.message || err) });
                        });
                    } catch (e) { /* ignore */ }
                }
                return promise;
            };
        } catch (e) { /* ignore */ }

        try {
            win.XMLHttpRequest.prototype.open = function (method, url) {
                this.__fifTapUrl = url;
                this.__fifTapMethod = method;
                return tap.origOpen.apply(this, arguments);
            };
            win.XMLHttpRequest.prototype.send = function (body) {
                const url = String(this.__fifTapUrl || '');
                if (netTapShouldLog(url)) {
                    this.addEventListener('load', () => {
                        try {
                            tap.upload({
                                via: 'xhr', url: url.slice(0, 400), method: this.__fifTapMethod,
                                req: body ? String(body).slice(0, 1500) : '',
                                status: this.status, res: String(this.responseText || '').slice(0, 20000)
                            });
                        } catch (e) { /* ignore */ }
                    });
                }
                return tap.origSend.apply(this, arguments);
            };
        } catch (e) { /* ignore */ }
    }

    let lastTapRecords = [];

    function stopNetTap() {
        const tap = netTapState;
        if (!tap) return;
        lastTapRecords = tap.records.slice(-80);
        try { pageWindow.fetch = tap.origFetch; } catch (e) { /* ignore */ }
        try {
            pageWindow.XMLHttpRequest.prototype.open = tap.origOpen;
            pageWindow.XMLHttpRequest.prototype.send = tap.origSend;
        } catch (e) { /* ignore */ }
        netTapState = null;
    }

    function extractTapParam(name) {
        const recs = (netTapState && netTapState.records) || lastTapRecords || [];
        for (let i = recs.length - 1; i >= 0; i--) {
            try {
                const u = new URL(recs[i].url);
                const v = u.searchParams.get(name);
                if (v) return v;
            } catch (e) { /* ignore */ }
        }
        return '';
    }

    // 练习拿不到题目时的接口排查：直接调用平台接口，对比不同参数下的返回
    async function probePracticeApi() {
        const activityId = extractTapParam('activityId');
        const courseId = extractTapParam('courseId');
        const batchNo = extractTapParam('batchNo');
        if (!activityId || !courseId) {
            uploadDiag({ tag: 'probe', error: 'no activityId/courseId', params: { activityId, courseId, batchNo } });
            return;
        }
        const base = 'https://gway.fifedu.com/istp-study';
        const q = `activityId=${activityId}&courseId=${courseId}&batchNo=${batchNo}`;
        const calls = [
            ['answerTopicHistory', `/questionApi/answerTopicHistory?${q}`],
            ['answerProgress', `/questionApi/answerProgress?${q}`],
            ['getTopicAnswerList', `/questionApi/getTopicAnswerList?batchNo=${batchNo}`],
            ['homePage', `/activityApi/getActivityHomePageData?activityId=${activityId}&courseId=${courseId}`],
            ['nextForce', `/questionApi/next?${q}&questionId=&ifForce=1&type=2&packageId=`],
            ['nextType1', `/questionApi/next?${q}&questionId=&ifForce=0&type=1&packageId=`],
            ['getTopicList', `/questionApi/getTopicList?${q}`]
        ];
        for (const item of calls) {
            const name = item[0];
            const uri = item[1];
            try {
                const res = await fetch(base + uri, { credentials: 'include' });
                const txt = await res.text();
                uploadDiag({
                    tag: 'probe', name,
                    ts: new Date().toLocaleTimeString('zh-CN', { hour12: false }),
                    url: base + uri, status: res.status, res: String(txt).slice(0, 6000)
                });
            } catch (e) {
                uploadDiag({ tag: 'probe', name, error: String((e && e.message) || e) });
            }
            await sleep(300);
        }
    }


    function describePageScripts() {
        try {
            const scripts = Array.from(document.querySelectorAll('script[src]')).map(s => String(s.src).slice(0, 300));
            const resources = (performance.getEntriesByType('resource') || [])
                .map(r => String(r.name))
                .filter(n => /\.js(\?|$)/.test(n) && /subject|practice|exercise|learn|chapter/i.test(n))
                .slice(0, 40);
            return { scripts: scripts.slice(0, 60), resources };
        } catch (e) {
            return { error: String(e && e.message || e) };
        }
    }

    function getChatEndpoint() {
        let url = String(state.config.aiBaseUrl || '').trim() || 'https://api.deepseek.com';
        if (!/\/chat\/completions\/?$/i.test(url)) {
            url = url.replace(/\/+$/, '') + '/chat/completions';
        }
        return url;
    }

    function gmHttpPost(url, headers, data, timeoutMs) {
        return new Promise((resolve, reject) => {
            if (typeof GM_xmlhttpRequest === 'function') {
                GM_xmlhttpRequest({
                    method: 'POST',
                    url,
                    headers,
                    data,
                    timeout: timeoutMs,
                    onload: res => resolve({ status: res.status, text: res.responseText || '' }),
                    onerror: () => reject(new Error('网络请求失败（请检查代理或网络）')),
                    ontimeout: () => reject(new Error('AI 接口请求超时'))
                });
                return;
            }
            fetch(url, { method: 'POST', headers, body: data })
                .then(async res => resolve({ status: res.status, text: await res.text() }))
                .catch(err => reject(err));
        });
    }

    // 统一的 AI 请求（含限速、429 退避、失败重试）
    async function requestAiChat(messages, maxTokens) {
        const key = String(state.config.aiApiKey || '').trim();
        if (!key) throw new Error('未配置 DeepSeek API Key');

        const gap = 1500 + Math.floor(Math.random() * 500);
        const wait = aiLastRequestAt + gap - Date.now();
        if (wait > 0) await sleep(wait);
        aiLastRequestAt = Date.now();

        const body = JSON.stringify({
            model: state.config.aiModel || 'deepseek-chat',
            messages,
            temperature: 0,
            max_tokens: maxTokens || 900,
            stream: false
        });
        const headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + key
        };

        let lastErr = null;
        for (let attempt = 0; attempt < 3; attempt++) {
            try {
                const res = await gmHttpPost(getChatEndpoint(), headers, body, 45000);
                if (res.status === 429 || res.status === 503) {
                    const backoff = 8000 + attempt * 7000 + Math.floor(Math.random() * 3000);
                    log(`AI 接口限流(${res.status})，${Math.round(backoff / 1000)} 秒后重试...`, 'warn');
                    await sleep(backoff);
                    continue;
                }
                if (res.status < 200 || res.status >= 300) {
                    let msg = res.text.slice(0, 180);
                    try {
                        const j = JSON.parse(res.text);
                        msg = (j.error && j.error.message) || j.message || msg;
                    } catch (e) { /* ignore */ }
                    throw new Error(`HTTP ${res.status}: ${msg}`);
                }
                const json = JSON.parse(res.text);
                const content = json.choices && json.choices[0] &&
                    json.choices[0].message && json.choices[0].message.content;
                if (!content) throw new Error('AI 返回内容为空');
                return String(content);
            } catch (err) {
                lastErr = err;
                if (attempt < 2) {
                    const backoff = 2500 + attempt * 2500;
                    log(`AI 请求失败，${Math.round(backoff / 1000)} 秒后重试：${err.message}`, 'warn');
                    await sleep(backoff);
                }
            }
        }
        throw lastErr || new Error('AI 请求失败');
    }

    function testAiConnection() {
        log(`正在测试 AI 接口（${state.config.aiModel || 'deepseek-chat'}）...`);
        return requestAiChat([
            { role: 'system', content: '你是接口连通性测试助手，只回复 OK。' },
            { role: 'user', content: 'ping' }
        ], 16).then(content => {
            log(`AI 接口可用：${String(content).slice(0, 60)}`, 'success');
            return true;
        }).catch(err => {
            log(`AI 接口测试失败：${err.message}`, 'error');
            return false;
        });
    }

    function detectQuestionKind(qType) {
        const t = Number(qType);
        if (t === 13) return '判断题';
        if (t === 1 || t === 12) return '多选题';
        return '单选题';
    }

    // 把 Vue 题目对象转换成 AI 可读的结构（支持阅读理解/完形填空的子题）
    function buildQuestionSet(question) {
        if (!question || question.id === undefined || question.id === null) return null;
        const subs = Array.isArray(question.subQuestionArray) ? question.subQuestionArray : [];
        const passage = subs.length ? htmlToPlainText(question.title) : '';
        const items = (subs.length ? subs : [question]).map(sub => ({
            stem: htmlToPlainText(sub.title),
            options: (Array.isArray(sub.answerOptionArray) ? sub.answerOptionArray : []).map(htmlToPlainText),
            qType: Number(sub.qType !== undefined ? sub.qType : question.qType),
            isComplex: Number(sub.isComplex !== undefined ? sub.isComplex : question.isComplex),
            raw: sub
        }));
        const optionTotal = items.reduce((n, it) => n + it.options.length, 0);
        if (!items.length || optionTotal < items.length * 2) return null;
        return { passage, items, raw: question };
    }

    function buildAiMessages(set) {
        const first = set.items[0] || {};
        const isComposite = set.items.length > 1 || !!set.passage;
        let typeName = detectQuestionKind(first.qType);
        if (isComposite) {
            typeName = (Number(first.isComplex) === 1 && [2, 10].includes(Number(first.qType)))
                ? '完形填空' : '阅读理解';
        }
        const payload = {
            type: typeName,
            passage: set.passage || undefined,
            questions: set.items.map(it => ({
                stem: it.stem,
                options: it.options.map((text, i) =>
                    `${String.fromCharCode(65 + i)}. ${String(text).replace(/^\s*[A-Za-z][.、)]\s*/, '')}`
                )
            }))
        };
        return [
            { role: 'system', content: String(state.config.aiPrompt || DEFAULT_AI_PROMPT) },
            { role: 'user', content: '请作答以下题目，严格按 JSON 返回：\n' + JSON.stringify(payload) }
        ];
    }

    function extractOptionLetters(value) {
        const s = String(value).toUpperCase();
        if (/正确|^对$|^√$|\bTRUE\b|^T$/.test(s)) return ['A'];
        if (/错误|^错$|^×$|\bFALSE\b|^F$/.test(s)) return ['B'];
        return s.match(/[A-Z]/g) || [];
    }

    function parseAiAnswerContent(content, set) {
        let text = String(content || '').trim();
        text = text.replace(/^```(?:json)?/i, '').replace(/```$/i, '').trim();
        const start = text.indexOf('{');
        const end = text.lastIndexOf('}');
        if (start >= 0 && end > start) text = text.slice(start, end + 1);

        const itemCount = set.items.length;
        let obj;
        try {
            obj = JSON.parse(text);
        } catch (e) {
            const m = text.match(/(?:答案|answer)\s*[:：]?\s*([A-Za-z,，、\s]+)/i);
            if (m) {
                obj = { answers: m[1].split(/[,，、\s]+/).filter(Boolean) };
            } else if (/^(正确|错误|对|错|√|×|true|false|t|f)$/i.test(text)) {
                // 兼容模型只回“正确/错误/√/×”等判断题答案
                obj = { answers: [text] };
            } else if (/^[\sA-Za-z,，、;；:：]+$/.test(text) && /[A-Za-z]/.test(text)) {
                // 兼容模型只回“B”“AC”“A,C”等纯字母答案
                const letters = text.match(/[A-Za-z]/g) || [];
                if (itemCount === 1) obj = { answers: [text] };
                else if (letters.length === itemCount) obj = { answers: letters };
                else obj = { answers: [text] };
            } else {
                throw new Error('AI 返回不是合法 JSON');
            }
        }

        let arr = obj.answers !== undefined ? obj.answers
            : obj.answer !== undefined ? obj.answer
            : obj.result !== undefined ? obj.result : obj.data;
        if (arr === undefined || arr === null) throw new Error('AI 返回缺少 answers 字段');
        if (!Array.isArray(arr)) arr = [arr];

        const count = set.items.length;
        let groups;
        if (count === 1) {
            groups = [arr];
        } else if (arr.length === count && arr.every(v => Array.isArray(v) || typeof v === 'string')) {
            groups = arr.map(v => (Array.isArray(v) ? v : [v]));
        } else {
            groups = [arr];
        }

        return groups.map((group, idx) => {
            const item = set.items[idx] || set.items[set.items.length - 1];
            const allowMultiple = [1, 12].includes(Number(item && item.qType));
            const letters = [];
            (Array.isArray(group) ? group : [group]).forEach(v => {
                extractOptionLetters(v).forEach(ch => letters.push(ch));
            });
            const unique = Array.from(new Set(letters));
            const picked = allowMultiple ? unique : unique.slice(0, 1);
            const maxOption = item && item.options ? item.options.length : 0;
            const indices = picked.map(ch => ch.charCodeAt(0) - 65).filter(i => i >= 0 && i < maxOption);
            // 不再兜底乱选；解析不出有效选项时返回空数组，交由上层重试或暂停
            return indices;
        });
    }

    function getAiCacheKey(set) {
        const signature = set.items.map(it =>
            `${it.qType}|${it.stem}|${it.options.join('~')}`
        ).join('||');
        // v2: 旧版本可能缓存了错误答案，换前缀使其失效
        return 'v2_' + hashString((set.passage || '') + '##' + signature);
    }

    function saveAiCache() {
        const keys = Object.keys(aiAnswerCache);
        if (keys.length > 300) {
            keys.sort((a, b) => (aiAnswerCache[a].t || 0) - (aiAnswerCache[b].t || 0));
            keys.slice(0, keys.length - 300).forEach(k => delete aiAnswerCache[k]);
        }
        Storage.set(AI_CACHE_KEY, aiAnswerCache);
    }

    let pendingAiCache = null;

    async function askAiForQuestionSet(set) {
        const cacheKey = getAiCacheKey(set);
        const cached = aiAnswerCache[cacheKey];
        if (cached && Array.isArray(cached.v) && cached.v.length === set.items.length &&
            cached.v.every(a => Array.isArray(a) && a.length > 0)) {
            log('命中本地答案缓存，跳过 AI 请求', 'info');
            return cached.v;
        }
        const content = await requestAiChat(buildAiMessages(set), 1000);
        log(`AI 原始回复：${String(content).replace(/\s+/g, ' ').slice(0, 120)}`, 'info');
        const answers = parseAiAnswerContent(content, set);
        const valid = Array.isArray(answers) && answers.length === set.items.length &&
            answers.every(a => Array.isArray(a) && a.length > 0);
        if (!valid) {
            throw new Error('AI 返回的答案不完整，已放弃本题以避免乱答');
        }
        // 先不写缓存，等平台判题确认答对后再缓存，避免把错误答案反复复用
        pendingAiCache = { cacheKey, answers };
        return answers;
    }

    function commitAiCache(correct) {
        if (correct && pendingAiCache) {
            aiAnswerCache[pendingAiCache.cacheKey] = { v: pendingAiCache.answers, t: Date.now() };
            saveAiCache();
        }
        pendingAiCache = null;
    }

    function applyAiAnswers(set, answers) {
        set.items.forEach((item, idx) => {
            const list = answers[idx];
            item.raw.checkVal = Array.isArray(list) ? list.slice() : [];
        });
    }

    // =========================================================================
    // 7.1 平台题目答案簿：从 /questionApi/next 响应提取标准答案（answerArray 为 0 基下标）
    // 平台在每次下发题目时会带回正确选项，比让 AI 猜更可靠
    // =========================================================================
    const platformAnswerBook = new Map();
    let lastCheckResult = null;

    function normQText(v) {
        return String(v || '')
            .replace(/<[^>]+>/g, '')
            .replace(/&nbsp;/g, ' ')
            .replace(/[\s\u00a0]+/g, '')
            .replace(/[（）()【】\[\]「」]/g, '');
    }

    function rememberQuestionFromResponse(json) {
        try {
            const data = (json && json.data) ? json.data : json;
            if (!data || data.questionId === undefined || data.questionId === null) return;
            const qid = String(data.questionId);
            if (!qid || qid === '-1' || qid === 'undefined') return;
            const subs = Array.isArray(data.subQuestionArray) ? data.subQuestionArray : [];
            const list = subs.length ? subs : [data];
            list.forEach(sub => {
                const answers = Array.isArray(sub.answerArray)
                    ? sub.answerArray.map(n => Number(n)).filter(n => Number.isFinite(n) && n >= 0)
                    : [];
                if (!answers.length) return;
                const rec = {
                    qid: String(sub.questionId !== undefined ? sub.questionId : qid),
                    qType: Number(sub.qType !== undefined ? sub.qType : data.qType),
                    title: normQText(sub.title),
                    options: (Array.isArray(sub.answerOptionArray) ? sub.answerOptionArray : []).map(normQText),
                    answers
                };
                if (rec.title) platformAnswerBook.set('t:' + rec.title, rec);
                if (rec.options.length) platformAnswerBook.set('o:' + rec.options.join('|'), rec);
                platformAnswerBook.set('id:' + rec.qid, rec);
            });
            while (platformAnswerBook.size > 400) {
                const firstKey = platformAnswerBook.keys().next().value;
                platformAnswerBook.delete(firstKey);
            }
        } catch (e) { /* ignore */ }
    }

    function inspectTapRecord(rec) {
        const url = String((rec && rec.url) || '');
        const text = String((rec && rec.res) || '');
        if (!text) return;
        if (/questionApi\/next/i.test(url)) {
            try { rememberQuestionFromResponse(JSON.parse(text)); } catch (e) { /* ignore */ }
        } else if (/questionApi\/check/i.test(url)) {
            try { lastCheckResult = JSON.parse(text); } catch (e) { lastCheckResult = null; }
            uploadDiag({
                tag: 'check',
                ts: new Date().toLocaleTimeString('zh-CN', { hour12: false }),
                res: text.slice(0, 1200)
            });
        }
    }

    // 读取指定容器内的练习题（纯 DOM，不依赖 Vue 实例）
    // 自适应练习的根节点是 .SubjectLib；测评弹窗的根节点是 .dialogBox
    function readQuestionDomFrom(root) {
        if (!root) return null;
        const reportEl = root.querySelector('.Report');
        if (reportEl && isVisibleEl(reportEl)) return { over: true, lib: root };

        let blocks = [];
        try { blocks = Array.from(root.querySelectorAll('.SubjectCon .RCJTmp')); } catch (e) { blocks = []; }
        if (blocks.length > 1) {
            // 只保留最外层题目容器，避免重复
            blocks = blocks.filter(b => !blocks.some(o => o !== b && b.contains(o)));
        }
        const items = blocks.map(block => {
            const contentEl = block.querySelector('.content');
            const title = normQText(contentEl ? contentEl.innerText : '');
            const options = Array.from(block.querySelectorAll('.RCJTmp-group-item')).map(it => ({
                el: it.querySelector('.myRCJTmp') || it,
                letter: normQText((it.querySelector('.left-option') || {}).innerText || ''),
                text: normQText(it.innerText || '')
            }));
            return { block, title, options };
        }).filter(it => it.options.length >= 2);

        if (!items.length) return null;
        return { over: false, lib: root, items };
    }

    function readPracticeQuestionDom() {
        let lib = null;
        try { lib = document.querySelector('.SubjectLib'); } catch (e) { lib = null; }
        if (!lib) return null;
        return readQuestionDomFrom(lib);
    }

    function practiceQuestionSignature(q) {
        return hashString(q.items.map(it =>
            it.title + '|' + it.options.map(o => o.text).join('~')
        ).join('##'));
    }

    function findPlatformAnswers(item) {
        const byTitle = platformAnswerBook.get('t:' + item.title);
        if (byTitle && byTitle.answers.length) return byTitle.answers.filter(i => i < item.options.length);
        const byOpts = platformAnswerBook.get('o:' + item.options.map(o => o.text).join('|'));
        if (byOpts && byOpts.answers.length) return byOpts.answers.filter(i => i < item.options.length);
        // 前缀/包含匹配：题干偶尔会被截断或带题型标签
        for (const rec of platformAnswerBook.values()) {
            if (!rec.title || !item.title) continue;
            const same = rec.title.indexOf(item.title) >= 0 || item.title.indexOf(rec.title) >= 0;
            if (same && rec.options.length === item.options.length) {
                const hit = rec.answers.filter(i => i < item.options.length);
                if (hit.length) return hit;
            }
        }
        return null;
    }

    async function resolveAnswersForDom(qItems) {
        const out = [];
        let needAi = false;
        qItems.forEach(item => {
            const ans = findPlatformAnswers(item);
            if (ans && ans.length) out.push(ans);
            else { out.push(null); needAi = true; }
        });
        if (!needAi) return out;
        if (!state.config.aiEnabled || !String(state.config.aiApiKey || '').trim()) return out;
        try {
            const set = {
                passage: '',
                items: qItems.map(it => ({
                    stem: it.title,
                    options: it.options.map(o => o.text),
                    qType: 9,
                    isComplex: 0,
                    raw: {}
                })),
                raw: {}
            };
            const aiAnswers = await askAiForQuestionSet(set);
            return out.map((a, i) => (a && a.length) ? a : (aiAnswers[i] || null));
        } catch (e) {
            log(`AI 兜底作答失败：${e.message}`, 'warn');
            return out;
        }
    }

    function fireHumanClick(el) {
        if (!el) return false;
        const rect = el.getBoundingClientRect ? el.getBoundingClientRect() : null;
        const x = rect ? Math.round(rect.left + rect.width / 2) : 0;
        const y = rect ? Math.round(rect.top + rect.height / 2) : 0;
        const base = { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0 };
        try {
            if (typeof PointerEvent === 'function') {
                el.dispatchEvent(new PointerEvent('pointerdown', base));
                el.dispatchEvent(new PointerEvent('pointerup', base));
            }
        } catch (e) { /* ignore */ }
        try {
            el.dispatchEvent(new MouseEvent('mousedown', base));
            el.dispatchEvent(new MouseEvent('mouseup', base));
        } catch (e) { /* ignore */ }
        try { el.click(); return true; } catch (e) { /* ignore */ }
        try { el.dispatchEvent(new MouseEvent('click', base)); return true; } catch (e) { return false; }
    }

    function findPracticeButton(pattern) {
        let lib = null;
        try { lib = document.querySelector('.SubjectLib'); } catch (e) { lib = null; }
        if (!lib) return null;
        const btns = Array.from(lib.querySelectorAll('.control button, button'));
        return btns.find(b => {
            if (!isVisibleEl(b) || b.disabled || b.classList.contains('is-disabled')) return false;
            return pattern.test(String(b.innerText || '').replace(/\s+/g, ''));
        }) || null;
    }

    // 读取练习进度（如「第 3 题 / 共 5 题」「3/5」），用于「报告页 + 进度」双信号完成判定
    function readPracticeProgress() {
        const lib = document.querySelector('.SubjectLib');
        if (!lib) return null;
        const text = String(lib.innerText || '').replace(/\s+/g, ' ');
        let m = text.match(/第\s*(\d+)\s*题\s*[\/／]?\s*共\s*(\d+)\s*题/);
        if (m) {
            const current = parseInt(m[1], 10);
            const total = parseInt(m[2], 10);
            if (total > 0 && current > 0 && current <= total) return { current, total };
        }
        m = text.match(/共\s*(\d+)\s*题[^0-9]{0,6}第\s*(\d+)\s*题/);
        if (m) {
            const total = parseInt(m[1], 10);
            const current = parseInt(m[2], 10);
            if (total > 0 && current > 0 && current <= total) return { current, total };
        }
        // 兜底：只认内容仅含「x/y」的独立进度元素，避免题干里的分数（如 1/2）被误判
        const progressEl = Array.from(lib.querySelectorAll('span, div, p, i, em, b')).find(el =>
            isVisibleEl(el) && /^\s*\d{1,3}\s*[\/／]\s*\d{1,3}\s*$/.test(String(el.innerText || '')));
        if (progressEl) {
            const parts = String(progressEl.innerText).trim().split(/[\/／]/).map(s => parseInt(s, 10));
            const current = parts[0], total = parts[1];
            if (total > 0 && total <= 200 && current > 0 && current <= total) return { current, total };
        }
        return null;
    }

    // 核心：纯 DOM 驱动答题（题目出现 -> 取标准答案 -> 勾选 -> 选完了 -> 下一题/报告）
    async function autoAnswerPracticeDom() {
        if (!acquireAiLock()) {
            log('检测到另一个标签页正在答题，已跳过，避免重复提交', 'warn');
            return { answered: 0, correct: 0, wrong: 0, completed: false };
        }
        const minDelay = clampNumber(state.config.aiMinDelayMs, 500, 30000, 2500);
        const maxDelay = Math.max(minDelay, clampNumber(state.config.aiMaxDelayMs, 500, 60000, 6000));
        const maxQuestions = clampNumber(state.config.aiMaxQuestions, 1, 200, 50);
        const doneSigs = new Set();
        let answered = 0, correct = 0, wrong = 0, idle = 0;
        let progressConfirmedDone = false;

        try {
            while (state.config.running && answered < maxQuestions) {
                if (aiLockHeldByOther()) {
                    log('另一个标签页已接管答题，本页退出以防重复提交', 'warn');
                    break;
                }
                touchAiLock();

                // 先读当前题目/报告状态：报告页优先结束，避免残留的「查看报告」按钮被反复点击
                const q = readPracticeQuestionDom();
                if (q && q.over) {
                    log('练习报告已显示，本节作答结束', 'info');
                    break;
                }

                if (!q) {
                    // 没有可读题目时，才允许点击「显示报告/查看报告」推进
                    const reportBtn = findPracticeButton(/显示报告|查看报告/);
                    if (reportBtn) {
                        fireHumanClick(reportBtn);
                        log('已点击【显示报告】', 'info');
                        await sleep(2000);
                        continue;
                    }
                    idle++;
                    if (idle > 40) {
                        log('等待题目超时，结束本节作答', 'warn');
                        break;
                    }
                    await sleep(900);
                    continue;
                }
                idle = 0;

                const sig = practiceQuestionSignature(q);
                if (doneSigs.has(sig)) {
                    // 双信号：进度已到最后一题且本题已提交，视为全部完成，不再继续点按钮
                    const progress = readPracticeProgress();
                    if (progress && progress.total > 0 && progress.current >= progress.total) {
                        progressConfirmedDone = true;
                        log(`双信号确认：进度 ${progress.current}/${progress.total}，本节题目已全部提交`, 'success');
                        const reportBtn = findPracticeButton(/显示报告|查看报告/);
                        if (reportBtn) {
                            fireHumanClick(reportBtn);
                            await sleep(1500);
                        }
                        break;
                    }
                    // 当前题已提交：优先点「下一题」继续后面的题，没有下一题才进入报告
                    const advanceBtn = findPracticeButton(/下一题/) ||
                        findPracticeButton(/显示报告|查看报告/);
                    if (advanceBtn) {
                        const label = (advanceBtn.innerText || '').trim();
                        fireHumanClick(advanceBtn);
                        log(`已点击【${label}】`, 'info');
                        await sleep(1800);
                        continue;
                    }
                    await sleep(800);
                    continue;
                }

                // AI 取答案失败时原地重试，避免单次抖动就带着未提交题目离开，触发「正在答题中」弹窗
                let answers = null;
                for (let attempt = 0; attempt < 3; attempt++) {
                    answers = await resolveAnswersForDom(q.items);
                    if (!answers.some(a => !a || !a.length)) break;
                    log(`未能取得本题可靠答案，第 ${attempt + 1} 次重试…`, 'warn');
                    await sleep(1500);
                }
                if (answers.some(a => !a || !a.length)) {
                    log('连续 3 次未能取得本题可靠答案，留在当前页等待下轮重试（请检查 AI 配置）', 'warn');
                    snapshotPracticeDOM('no-answer');
                    break;
                }

                let clicked = 0;
                q.items.forEach((item, qi) => {
                    answers[qi].forEach(idx => {
                        const opt = item.options[idx];
                        if (opt && fireHumanClick(opt.el)) clicked++;
                    });
                });
                const preview = answers.map(a => a.map(i => String.fromCharCode(65 + i)).join('')).join(' / ');
                log(`已选择答案 ${preview}（点击 ${clicked} 个选项），即将提交…`, 'info');

                const thinkMs = minDelay + Math.floor(Math.random() * (maxDelay - minDelay + 1));
                await sleep(thinkMs);
                if (!state.config.running) break;

                const submitBtn = findPracticeButton(/选完了|提交|确定/);
                if (!submitBtn) {
                    log('未找到【选完了】提交按钮，等待平台自动处理…', 'warn');
                    await sleep(1500);
                    continue;
                }
                lastCheckResult = null;
                fireHumanClick(submitBtn);
                doneSigs.add(sig);
                answered++;
                log(`第 ${answered} 题已提交，等待判题…`, 'info');

                // 等待判题结果或自动跳题
                for (let i = 0; i < 24 && state.config.running; i++) {
                    await sleep(600);
                    const now = readPracticeQuestionDom();
                    if (!now || now.over) break;
                    if (practiceQuestionSignature(now) !== sig) break;
                }

                const chk = lastCheckResult;
                if (chk) {
                    const d = (chk && chk.data) ? chk.data : chk;
                    const right = !!(d && (d.isCorrect === 1 || d.isCorrect === true ||
                        d.correct === 1 || d.correct === true));
                    if (right) correct++; else wrong++;
                }
            }
        } finally {
            releaseAiLock();
        }

        const over = !!(document.querySelector('.SubjectLib .Report'));
        const completed = over || progressConfirmedDone;
        if (completed) {
            const progress = readPracticeProgress();
            const suffix = progress ? `（进度 ${progress.current}/${progress.total}）` : '';
            log(`本节练习作答完成${suffix}：提交 ${answered} 题，正确 ${correct}，错误 ${wrong}`, 'success');
        }
        return { answered, correct, wrong, completed };
    }

    // =========================================================================
    // 7.2 课程「测评评估掌握度」自动答题
    // 入口：课程首页掌握度卡片 .knowledge-mastery-desc .inner_text
    // 弹窗：HappyDrink 封面 -> .dialogBox 答题 -> .ReportBox 报告
    // =========================================================================
    function findEvaluationTrigger() {
        const desc = document.querySelector('.knowledge-mastery-desc');
        if (desc) {
            // 优先命中平台渲染出的文字入口（含脚本补全的入口）
            const textEntry = Array.from(desc.querySelectorAll('.inner_text, span')).find(el =>
                isVisibleEl(el) && /测评评估掌握度/.test(String(el.innerText || '').trim()));
            if (textEntry) return textEntry;
            // 首次进入时平台可能因异步数据竞态没有渲染文字，外层 span 本身绑定了点击事件，可直接作为触发点
            const clickable = Array.from(desc.children).find(el =>
                el.tagName === 'SPAN' && getComputedStyle(el).display !== 'none');
            if (clickable) return clickable;
        }
        return Array.from(document.querySelectorAll('.knowledge-mastery-desc .inner_text, .knowledge-mastery-desc span'))
            .find(el => isVisibleEl(el) && /^测评评估掌握度$/.test(String(el.innerText || '').trim())) || null;
    }

    // 首次进入课程首页时，平台掌握度卡片依赖异步接口（checkIsReply / GetKnowledgePointCollect），
    // 数据未就绪时「测评评估掌握度」入口不会渲染，且不会自动重试。这里做入口强制渲染校验与补全。
    let evalEntryGuardStarted = false;
    function ensureEvaluationEntry() {
        const repair = () => {
            const desc = document.querySelector('.knowledge-mastery-desc');
            if (!desc) return;
            const real = Array.from(desc.querySelectorAll('.inner_text:not([data-fif-inject])'))
                .find(el => /测评评估掌握度/.test(String(el.innerText || '').trim()));
            const mine = desc.querySelector('.inner_text[data-fif-inject]');
            if (real) {
                if (mine) mine.remove();
                delete desc.dataset.fifEntryWait;
                return;
            }
            if (mine) return;
            if (!desc.dataset.fifEntryWait) {
                desc.dataset.fifEntryWait = String(Date.now());
                return;
            }
            if (Date.now() - Number(desc.dataset.fifEntryWait) < 1800) return;

            const host = Array.from(desc.children).find(el => el.tagName === 'SPAN') || desc;
            const span = document.createElement('span');
            span.className = 'inner_text';
            span.setAttribute('data-fif-inject', '1');
            span.textContent = '测评评估掌握度';
            span.style.cssText = 'display:inline-block;height:28px;line-height:28px;font-size:12px;color:#0a84ff;cursor:pointer;';
            const scopeAttr = Array.from(desc.attributes).find(a => /^data-v-/.test(a.name));
            if (scopeAttr) span.setAttribute(scopeAttr.name, scopeAttr.value);
            host.appendChild(span);
            log('【测评评估掌握度】入口未渲染，已自动补全（无需刷新页面）', 'success');
        };
        repair();
        if (evalEntryGuardStarted) return;
        evalEntryGuardStarted = true;
        setInterval(repair, 900);
        window.addEventListener('resize', repair);
    }

    function evaluationDialogEl() {
        const dlg = Array.from(document.querySelectorAll('.CEDialog')).find(isVisibleEl);
        if (dlg) return dlg;
        // 兼容自定义 class 未落到 .el-dialog 的情况：封面 .HappyDrink 出现即视为弹窗已打开
        return Array.from(document.querySelectorAll('.HappyDrink')).find(isVisibleEl) || null;
    }

    function evaluationAnsweringEl() {
        return Array.from(document.querySelectorAll('.dialogBox')).find(isVisibleEl) || null;
    }

    function evaluationMessageText() {
        const els = Array.from(document.querySelectorAll('.el-message-box__message, .el-message'));
        for (const el of els) {
            if (!isVisibleEl(el)) continue;
            const text = String(el.innerText || '').replace(/\s+/g, ' ').trim();
            if (text) return text;
        }
        return '';
    }

    function isEvaluationBlockedMessage(msg) {
        return /时间过短|暂无可测评|已经掌握|最短时间间隔/.test(String(msg || ''));
    }

    // 「暂无可测评题目」单独识别：可通过强制重选学习目标刷新题池，不应直接终止
    function isNoTopicMessage(msg) {
        return /暂无可测评|无可测评题目|暂无题目/.test(String(msg || ''));
    }

    function closeEvaluationMessageBox() {
        const btn = Array.from(document.querySelectorAll('.el-message-box__btns button, .el-message-box__headerbtn'))
            .find(isVisibleEl);
        if (btn) { fireHumanClick(btn); return true; }
        return false;
    }

    function findEvaluationButton(pattern, root) {
        const scope = root || Array.from(document.querySelectorAll('.dialogBox')).find(isVisibleEl);
        if (!scope) return null;
        return Array.from(scope.querySelectorAll('button, .el-button')).find(btn => {
            if (!isVisibleEl(btn) || btn.disabled || btn.classList.contains('is-disabled')) return false;
            return pattern.test(String(btn.innerText || '').replace(/\s+/g, ''));
        }) || null;
    }

    function readEvaluationQuestionDom() {
        const boxes = Array.from(document.querySelectorAll('.dialogBox')).filter(isVisibleEl);
        for (const box of boxes) {
            const q = readQuestionDomFrom(box);
            if (q) return q;
        }
        return null;
    }

    function isEvaluationReportVisible() {
        return Array.from(document.querySelectorAll('.ReportBox'))
            .some(box => isVisibleEl(box) && !!box.querySelector('.Report'));
    }

    function findReportButton(pattern) {
        const boxes = Array.from(document.querySelectorAll('.ReportBox')).filter(isVisibleEl);
        for (const box of boxes) {
            const btn = Array.from(box.querySelectorAll('button, .el-button')).find(b =>
                isVisibleEl(b) && !b.disabled && pattern.test(String(b.innerText || '').trim()));
            if (btn) return btn;
        }
        return null;
    }

    // 测评前置校验与补全：首次测评时学习目标为空（覆盖显示「课程默认」），直接点「开始测评」会被平台拒绝
    // （verifyPushTopic 返回 isNextQue=false → 暂无可测评题目）。这里自动完成「更改 → 全选 → 确定」。
    async function prepareEvaluationScope(dialog, force) {
        if (!dialog) return false;
        const scopeText = () => String((dialog.querySelector('.learningTarget .num') || {}).innerText || '').replace(/\s+/g, '');
        const hasScope = () => {
            const t = scopeText();
            return /\d+个/.test(t) && !/^0个/.test(t);
        };
        if (!force && hasScope()) {
            log(`测评覆盖范围已就绪：${scopeText()}`, 'info');
            return true;
        }
        log(force
            ? `强制重新选择学习目标（当前覆盖「${scopeText() || '课程默认'}」），刷新可测评题目…`
            : `测评覆盖范围为「${scopeText() || '课程默认'}」，正在自动选择学习目标…`, 'info');

        const changeBtn = dialog.querySelector('.learningTarget .btn');
        if (!changeBtn) {
            log('未找到【更改】按钮，无法补全学习目标', 'warn');
            return false;
        }
        fireHumanClick(changeBtn);

        let setDlg = null;
        const findSetDialog = () => Array.from(document.querySelectorAll('.set-target')).find(el => {
            const box = el.querySelector('.set-box') || el;
            return isVisibleEl(el) || isVisibleEl(box);
        }) || null;
        for (let i = 0; i < 25 && !setDlg; i++) {
            setDlg = findSetDialog();
            if (!setDlg) await sleep(400);
        }
        if (!setDlg) {
            log('【设置学习目标】弹窗未打开，已中止本次测评', 'warn');
            return false;
        }
        for (let i = 0; i < 30; i++) {
            if (setDlg.querySelector('.el-tree-node')) break;
            await sleep(400);
        }

        const findSelectAll = () => setDlg.querySelector('.set-left .set-content-title .link') ||
            Array.from(setDlg.querySelectorAll('.set-left .link')).find(el => /全选/.test(String(el.innerText || '')));
        const readPicked = () => {
            const t = String((setDlg.querySelector('.set-right .set-content-title') || {}).innerText || '');
            const m = t.match(/已选[：:]\s*(\d+)/);
            return m ? parseInt(m[1], 10) : 0;
        };
        const waitPicked = async (ok, times, gap) => {
            for (let i = 0; i < times; i++) {
                if (ok(readPicked())) return true;
                await sleep(gap);
            }
            return ok(readPicked());
        };

        let selectAll = findSelectAll();
        if (!selectAll) {
            log('未找到【全选】按钮，已中止本次测评', 'warn');
            return false;
        }
        if (force && /取消全选/.test(String(selectAll.innerText || ''))) {
            // 已全选状态：先取消全选再重新全选，让学习目标发生真实变更，强制平台刷新可测题池
            log('当前已是全选，先【取消全选】再重新【全选】以刷新题池…', 'info');
            fireHumanClick(selectAll);
            // 必须等「已选 N 个」真正归零并留出渲染时间，否则平台的勾选状态还没提交就点全选，会变成空操作
            const cleared = await waitPicked(n => n === 0, 30, 400);
            if (!cleared) log(`取消全选后已选数量未归零（${readPicked()}），仍继续尝试重新全选`, 'warn');
            await sleep(1200 + Math.floor(Math.random() * 800));
            selectAll = findSelectAll();
            if (!selectAll) {
                log('取消全选后未找到【全选】按钮，已中止本次测评', 'warn');
                return false;
            }
        }

        // 勾选全选：点击后已选数量若仍为 0，说明 Vue 还没渲染完或点击未生效，重试点击
        let picked = readPicked();
        for (let attempt = 0; attempt < 4 && picked === 0; attempt++) {
            selectAll = findSelectAll();
            if (!selectAll) break;
            if (!/取消全选/.test(String(selectAll.innerText || ''))) {
                fireHumanClick(selectAll);
                if (attempt === 0) log('已点击【全选】，等待知识点勾选完成…', 'info');
            }
            await sleep(900 + Math.floor(Math.random() * 500));
            await waitPicked(n => n > 0, 12, 400);
            picked = readPicked();
            if (!picked && attempt < 3) {
                log(`全选后仍未选中知识点（第 ${attempt + 1} 次尝试），稍后重试…`, 'warn');
                await sleep(1000);
            }
        }
        if (!picked) {
            log('全选后仍未选中任何知识点，已中止本次测评', 'warn');
            return false;
        }
        log(`已勾选 ${picked} 个学习目标，正在保存…`, 'info');

        const okBtn = setDlg.querySelector('.set-right .set-ensure') ||
            Array.from(setDlg.querySelectorAll('.set-ensure')).find(el => /^确定$/.test(String(el.innerText || '').trim()));
        if (!okBtn) {
            log('未找到【确定】按钮，已中止本次测评', 'warn');
            return false;
        }
        fireHumanClick(okBtn);

        for (let i = 0; i < 40; i++) {
            await sleep(500);
            if (!findSetDialog()) break;
        }
        for (let i = 0; i < 30; i++) {
            if (hasScope()) break;
            await sleep(500);
        }
        const ok = hasScope();
        if (ok) await sleep(800 + Math.floor(Math.random() * 500));
        log(ok ? `学习目标已保存，测评覆盖：${scopeText()}` : `学习目标保存后覆盖范围仍为「${scopeText() || '课程默认'}」`,
            ok ? 'success' : 'warn');
        return ok;
    }

    async function autoAnswerEvaluationDom() {
        if (!acquireAiLock()) {
            log('检测到另一个标签页正在答题，已跳过测评，避免重复提交', 'warn');
            return { answered: 0, correct: 0, wrong: 0, completed: false };
        }
        const minDelay = clampNumber(state.config.aiMinDelayMs, 500, 30000, 2500);
        const maxDelay = Math.max(minDelay, clampNumber(state.config.aiMaxDelayMs, 500, 60000, 6000));
        const maxQuestions = clampNumber(state.config.aiMaxQuestions, 1, 200, 50);
        const doneSigs = new Set();
        let answered = 0, correct = 0, wrong = 0, idle = 0, stuckSig = '', stuckCount = 0;

        try {
            while (state.config.running && answered < maxQuestions) {
                if (aiLockHeldByOther()) {
                    log('另一个标签页已接管答题，本页退出以防重复提交', 'warn');
                    break;
                }
                touchAiLock();

                if (isEvaluationReportVisible()) break;

                const reportBtn = findEvaluationButton(/^显示报告$/);
                if (reportBtn) {
                    fireHumanClick(reportBtn);
                    log('测评已全部作答，点击【显示报告】', 'info');
                    await sleep(2500);
                    continue;
                }

                const q = readEvaluationQuestionDom();
                if (!q || q.over) {
                    idle++;
                    if (idle > 40) {
                        log('等待测评题目超时，结束本次测评作答', 'warn');
                        break;
                    }
                    await sleep(900);
                    continue;
                }
                idle = 0;

                const sig = practiceQuestionSignature(q);
                if (sig === stuckSig) stuckCount++; else { stuckSig = sig; stuckCount = 0; }
                if (stuckCount > 100) {
                    log('测评题目长时间未变化，已停止本次作答', 'error');
                    break;
                }

                if (doneSigs.has(sig)) {
                    const nextBtn = findEvaluationButton(/^下一题$/);
                    if (nextBtn) {
                        fireHumanClick(nextBtn);
                        log('已点击【下一题】', 'info');
                        await sleep(1200);
                    } else {
                        await sleep(1200);
                    }
                    continue;
                }

                // 稍等平台 next 响应写入答案簿，避免题面先渲染、答案后到达的竞态
                await sleep(700);
                const answers = await resolveAnswersForDom(q.items);
                if (answers.some(a => !a || !a.length)) {
                    log('未能取得本题可靠答案，已暂停测评作答（平台未返回标准答案且 AI 兜底不可用）', 'warn');
                    break;
                }

                let clicked = 0;
                q.items.forEach((item, qi) => {
                    answers[qi].forEach(idx => {
                        const opt = item.options[idx];
                        if (opt && fireHumanClick(opt.el)) clicked++;
                    });
                });
                const preview = answers.map(a => a.map(i => String.fromCharCode(65 + i)).join('')).join(' / ');
                log(`测评已选择答案 ${preview}（点击 ${clicked} 个选项），即将提交…`, 'info');

                const thinkMs = minDelay + Math.floor(Math.random() * (maxDelay - minDelay + 1));
                await sleep(thinkMs);
                if (!state.config.running) break;

                const submitBtn = findEvaluationButton(/^选完了$/);
                if (!submitBtn) {
                    log('未找到【选完了】提交按钮，结束本次测评作答', 'warn');
                    break;
                }
                lastCheckResult = null;
                fireHumanClick(submitBtn);
                doneSigs.add(sig);
                answered++;
                log(`测评第 ${answered} 题已提交，等待判题…`, 'info');

                for (let i = 0; i < 24 && state.config.running; i++) {
                    await sleep(600);
                    if (isEvaluationReportVisible()) break;
                    if (findEvaluationButton(/^下一题$/) || findEvaluationButton(/^显示报告$/)) break;
                    const now = readEvaluationQuestionDom();
                    if (now && !now.over && practiceQuestionSignature(now) !== sig) break;
                }

                const chk = lastCheckResult;
                const d = (chk && chk.data) ? chk.data : chk;
                const nextBtn = findEvaluationButton(/^下一题$/);
                const right = d
                    ? !!(d.status === 1 || d.isCorrect === 1 || d.isCorrect === true ||
                        d.correct === 1 || d.correct === true)
                    : !nextBtn;
                if (right) correct++; else wrong++;

                // 答错时平台会停留在本题并给出「下一题」
                if (nextBtn) {
                    fireHumanClick(nextBtn);
                    log('已点击【下一题】', 'info');
                    await sleep(1200);
                }
            }
        } finally {
            releaseAiLock();
        }

        const completed = isEvaluationReportVisible();
        if (completed) log(`测评作答完成：提交 ${answered} 题，正确 ${correct}，错误 ${wrong}`, 'success');
        return { answered, correct, wrong, completed };
    }

    // 目标熟练度输入解析：选填；一旦填写则必须是 0-100 的整数
    function parseEvalTargetInput(raw) {
        const s = String(raw == null ? '' : raw).trim();
        if (!s) return { provided: false, valid: true, value: null, message: '' };
        if (!/^\d+$/.test(s)) return { provided: true, valid: false, value: null, message: '目标熟练度需为 0-100 的整数' };
        const n = Number(s);
        if (!Number.isInteger(n) || n < 0 || n > 100) {
            return { provided: true, valid: false, value: null, message: '目标熟练度需在 0-100 之间（整数）' };
        }
        return { provided: true, valid: true, value: n, message: '' };
    }

    // 循环判定：只有「开关开启 + 目标合法 + 本次熟练度低于目标」才继续下一轮
    function shouldContinueEvalLoop(loopEnabled, target, mastery) {
        if (!loopEnabled) return false;
        if (target == null || mastery == null) return false;
        const t = Number(target), m = Number(mastery);
        if (!Number.isFinite(t) || !Number.isFinite(m)) return false;
        return m < t;
    }

    function getEvaluationLoopConfig() {
        const target = parseEvalTargetInput(state.config.evalTargetMastery);
        const infinite = !!state.config.evalInfinite;
        return {
            enabled: infinite || !!state.config.evalLoop,
            infinite,
            target,
            // 无限循环模式忽略目标值；普通循环模式需要开关开启且目标已填、合法
            active: infinite || (!!state.config.evalLoop && target.provided && target.valid)
        };
    }

    function evalStopRequested() {
        return !!state.evalStopRequested || !state.config.running;
    }

    // 从测评报告中读取本次熟练度（掌握度）数值，读取失败返回 null
    function extractEvaluationMastery() {
        const report = Array.from(document.querySelectorAll('.ReportBox')).find(isVisibleEl);
        const dialog = evaluationDialogEl();
        const scopes = [];
        if (report) scopes.push(report);
        if (dialog && dialog !== report) scopes.push(dialog);

        // 报告中的掌握度仪表盘：.gaugeCon > .textCon > .n = 数值%，.b = 「掌握度」
        for (const scope of scopes) {
            const gauges = Array.from(scope.querySelectorAll('.gaugeCon'));
            for (const g of gauges) {
                const label = String((g.querySelector('.textCon .b') || {}).innerText || '').trim();
                if (label && !/掌握度/.test(label)) continue;
                const txt = String((g.querySelector('.textCon .n') || {}).innerText || g.innerText || '');
                const m = txt.match(/(\d+(?:\.\d+)?)\s*%/);
                if (m) return Number(m[1]);
            }
        }
        // 兜底：报告文本里的「掌握度 xx%」（排除「掌握度变化」）
        const text = report ? String(report.innerText || '') : '';
        const m2 = text.match(/掌握度(?!变化)[^0-9%]{0,12}(\d+(?:\.\d+)?)\s*%/);
        if (m2) return Number(m2[1]);
        // 再兜底：课程首页掌握度卡片
        const home = document.querySelector('.knowledge-mastery');
        if (home) {
            const m3 = String(home.innerText || '').match(/(\d+(?:\.\d+)?)\s*%/);
            if (m3) return Number(m3[1]);
        }
        return null;
    }

    // 关闭测评报告并返回课程首页（与原有单轮测评的收尾动作一致）
    async function closeEvaluationReport() {
        await sleep(1500);
        const enterBtn = findReportButton(/^进入学习$/);
        if (enterBtn) {
            fireHumanClick(enterBtn);
            log('已关闭测评报告，返回课程首页', 'info');
            await sleep(1200);
        }
    }

    // 跳转前等待测评界面完全退出（答题弹窗/报告/确认框都消失），降低触发「正在答题中」弹窗的概率
    async function waitEvaluationUiClosed(timeoutMs = 10000) {
        const start = Date.now();
        while (Date.now() - start < timeoutMs) {
            if (!evaluationAnsweringEl() && !isEvaluationReportVisible() && !visibleModals().length) {
                return true;
            }
            await sleep(500);
        }
        log('测评界面在超时时间内仍未完全关闭，尝试清理后再跳转', 'warn');
        await cleanupEvaluationUi('跳转前仍有残留界面');
        return false;
    }

    async function runCourseEvaluation({ singleRound = false, maxRounds: maxRoundsArg } = {}) {
        if (state.evaluationRunning) return;
        ensureEvaluationEntry();
        // 已配置 AI Key 时自动启用 AI 作答，保证测评启动后可自动完成全部题目
        if (!state.config.aiEnabled && String(state.config.aiApiKey || '').trim()) {
            state.config.aiEnabled = true;
            log('检测到已配置 AI Key，已自动启用 AI 作答', 'info');
        }
        if (!String(state.config.aiApiKey || '').trim()) {
            log('未配置 AI API Key：优先使用平台下发的标准答案，平台未下发答案的题目将暂停作答', 'warn');
        }

        let loopCfg = getEvaluationLoopConfig();
        if (singleRound) {
            // 一键达标模式：每门课只测一轮，测完立即离开，避免多轮测评后跳转触发「正在答题中」弹窗
            loopCfg = { ...loopCfg, enabled: false, active: false, infinite: false };
            log('【一键达标】本次测评按单轮执行：不做达标检测与自动重测', 'info');
        }
        if (!loopCfg.infinite && loopCfg.enabled && loopCfg.target.provided && !loopCfg.target.valid) {
            log(`目标熟练度指标无效：${loopCfg.target.message}，请修正后再开始测评`, 'warn');
            updateEvalTargetHint();
            return;
        }
        if (!loopCfg.infinite && loopCfg.enabled && !loopCfg.target.provided) {
            log('循环测试已开启，但未填写目标熟练度指标：本次按单轮测评执行', 'info');
        }
        if (loopCfg.infinite) {
            log('【高风险】无限循环已开启：忽略目标熟练度，每轮结束自动开始下一轮完整测评，只有手动点击「停止测评」或出现平台拦截/连续失败才会终止', 'warn');
        } else if (loopCfg.active) {
            log(`循环测试已开启：目标熟练度 ${loopCfg.target.value}%，未达到将自动开始新一轮完整测评；达到目标、手动停止或出现异常时立即终止`, 'info');
        }

        state.evaluationRunning = true;
        state.evalStopRequested = false;
        const prevRunning = state.config.running;
        state.config.running = true;
        updatePanelUI();

        // 轮次上限：调用方可指定（如一键达标阶段 B）；否则按模式取默认（无限 500 / 普通 20）
        const defaultMaxRounds = loopCfg.infinite ? 500 : 20;
        const maxRounds = (Number.isFinite(maxRoundsArg) && maxRoundsArg > 0) ? Math.floor(maxRoundsArg) : defaultMaxRounds;
        const retryableReasons = ['no-topic-after-retry', 'questions-not-ready', 'dialog-not-open', 'no-start-button'];
        let roundNo = 0;
        let softFailures = 0;
        try {
            while (true) {
                roundNo++;
                if (evalStopRequested()) {
                    log('已手动停止测评，终止流程', 'warn');
                    break;
                }
                if (loopCfg.active) {
                    log(loopCfg.infinite ? `—— 开始第 ${roundNo} 轮测评（无限循环）——` : `—— 开始第 ${roundNo} 轮测评 ——`, 'info');
                }

                let round;
                try {
                    round = await runSingleEvaluationRound();
                } catch (err) {
                    log(`第 ${roundNo} 轮测评出现异常：${err && err.message ? err.message : err}，已终止循环`, 'error');
                    break;
                }

                if (!round || round.status !== 'completed') {
                    const reason = (round && round.reason) || '未知原因';
                    if (reason === 'manual-stop') {
                        log('已手动停止测评，终止流程', 'warn');
                        break;
                    }
                    if (!loopCfg.active) break;
                    // 题池/页面未就绪属于可自愈问题：关闭残留报告、等待后重试，避免循环被一次偶发失败打断
                    if (retryableReasons.includes(reason) && softFailures < 2) {
                        softFailures++;
                        const waitSec = 6 + softFailures * 4;
                        log(`第 ${roundNo} 轮测评未完成（${reason}），${waitSec} 秒后自动重试（第 ${softFailures}/2 次）`, 'warn');
                        await closeEvaluationReport();
                        await sleep(waitSec * 1000);
                        continue;
                    }
                    log(`第 ${roundNo} 轮测评未完成（${reason}），终止循环`, 'warn');
                    break;
                }
                softFailures = 0;

                const mastery = round.mastery;
                if (!loopCfg.active) {
                    // 单轮模式：与原有版本完全一致，看完报告后返回课程首页
                    await closeEvaluationReport();
                    break;
                }

                if (loopCfg.infinite) {
                    log(mastery == null
                        ? `第 ${roundNo} 轮测评已完成（未能读取熟练度数值，继续下一轮）`
                        : `第 ${roundNo} 轮测评熟练度：${mastery}%（无限循环模式，忽略目标值）`, 'info');
                } else {
                    if (mastery == null) {
                        log('未能读取本次测评熟练度数值，为避免无限重复已终止循环', 'warn');
                        await closeEvaluationReport();
                        break;
                    }
                    log(`第 ${roundNo} 轮测评熟练度：${mastery}%（目标 ${loopCfg.target.value}%）`, 'info');
                    if (!shouldContinueEvalLoop(true, loopCfg.target.value, mastery)) {
                        log(`熟练度已达目标（${mastery}% ≥ ${loopCfg.target.value}%），循环测试结束`, 'success');
                        await closeEvaluationReport();
                        break;
                    }
                }
                if (evalStopRequested()) {
                    log('已手动停止测评，终止循环', 'warn');
                    break;
                }
                if (roundNo >= maxRounds) {
                    log(loopCfg.infinite
                        ? `已连续测评 ${maxRounds} 轮，达到安全上限，已终止循环`
                        : `已连续测评 ${maxRounds} 轮仍未达到目标，为避免过度请求已终止循环`, 'warn');
                    await closeEvaluationReport();
                    break;
                }

                log(loopCfg.infinite
                    ? `准备开始第 ${roundNo + 1} 轮完整测评（学习目标校验 → 开始测评 → 自动答题 → 查看报告）`
                    : `本次熟练度 ${mastery}% 低于目标 ${loopCfg.target.value}%，将自动开始第 ${roundNo + 1} 轮完整测评（学习目标校验 → 开始测评 → 自动答题 → 查看报告）`, 'info');

                // 「再测一次」按钮可能因报告渲染延迟而暂时找不到，这里多等几秒
                let againBtn = findReportButton(/^再测一次$/);
                for (let i = 0; i < 10 && !againBtn; i++) {
                    await sleep(600);
                    againBtn = findReportButton(/^再测一次$/);
                }
                if (!againBtn) {
                    // 兜底：关闭报告后回到课程首页，由下一轮重新从测评入口启动，不再直接终止循环
                    log('未找到【再测一次】按钮，关闭报告后从测评入口重新开始下一轮', 'warn');
                    await closeEvaluationReport();
                    await sleep(1500);
                    continue;
                }
                fireHumanClick(againBtn);
                // 等待平台从报告页切回测评封面，再进入下一轮完整流程
                let backToCover = false;
                for (let i = 0; i < 30; i++) {
                    await sleep(600);
                    if (!isEvaluationReportVisible()) { backToCover = true; break; }
                }
                if (!backToCover) log('点击【再测一次】后测评报告未关闭，稍后继续尝试下一轮', 'warn');
                // 轮次之间留出随机冷却，降低连续测评被平台判定为高频的概率
                await sleep(loopCfg.infinite
                    ? 4000 + Math.floor(Math.random() * 3000)
                    : 900 + Math.floor(Math.random() * 700));
            }
        } finally {
            const stoppedByUser = state.evalStopRequested;
            // 无论正常结束、异常、手动停止还是轮次上限，都清扫测评残留弹窗，避免卡住主循环
            await cleanupEvaluationUi(stoppedByUser ? '测评已手动停止' : '测评流程结束');
            state.config.running = stoppedByUser ? false : prevRunning;
            state.evaluationRunning = false;
            state.evalStopRequested = false;
            saveConfig();
            updatePanelUI();
        }
    }

    // 单轮完整测评：打开入口 → 学习目标校验补全 → 开始测评 → 自动答题 → 读取报告
    // 循环模式下每一轮都完整执行本函数，保证与原有单轮流程、交互逻辑完全一致
    async function runSingleEvaluationRound() {
        if (evaluationDialogEl() || evaluationAnsweringEl()) {
            log('测评弹窗已经打开，直接开始自动答题', 'info');
        } else {
            let trigger = findEvaluationTrigger();
            for (let i = 0; i < 16 && !trigger; i++) {
                await sleep(500);
                trigger = findEvaluationTrigger();
            }
            if (!trigger) {
                log('未找到【测评评估掌握度】入口：请先进入课程首页，并等掌握度卡片加载完成', 'warn');
                return { status: 'aborted', reason: 'no-entry' };
            }
            log('正在打开【测评评估掌握度】...', 'info');
            fireHumanClick(trigger);

            let dialog = null;
            for (let i = 0; i < 40; i++) {
                dialog = evaluationDialogEl();
                if (dialog) break;
                const msg = evaluationMessageText();
                if (isEvaluationBlockedMessage(msg)) {
                    log(`平台提示：${msg.slice(0, 160)}`, 'warn');
                    closeEvaluationMessageBox();
                    return { status: 'aborted', reason: 'blocked' };
                }
                await sleep(500);
            }
            if (!dialog) {
                log('测评弹窗未打开（可能距离上次测评时间过短或暂无可测评题目）', 'warn');
                return { status: 'aborted', reason: 'dialog-not-open' };
            }
        }

        if (evalStopRequested()) return { status: 'aborted', reason: 'manual-stop' };

        // 封面：先做覆盖范围前置校验与补全（修复必须手动「更改→全选→确定」才能测评的问题）
        if (!evaluationAnsweringEl() && !isEvaluationReportVisible()) {
            await prepareEvaluationScope(evaluationDialogEl());
        }

        // 封面：点击「开始测评」并等待题目/报告。
        // 平台在学习目标未变化时可能对旧题池返回「暂无可测评题目」，此时按手动可行路径
        // 强制重走「更改 → 取消全选 → 重新全选 → 确定」刷新题池后重试（最多 3 次）；其余拦截提示仍直接终止。
        let ready = false;
        for (let attempt = 0; attempt < 3 && !ready; attempt++) {
            if (evaluationAnsweringEl() || isEvaluationReportVisible()) { ready = true; break; }
            if (evalStopRequested()) return { status: 'aborted', reason: 'manual-stop' };
            if (attempt > 0) {
                // 上次失败后平台可能仍在收尾，先等一会再改学习目标
                await sleep(2000 + Math.floor(Math.random() * 1500));
                const scopeOk = await prepareEvaluationScope(evaluationDialogEl(), true);
                if (!scopeOk) log('重新选择学习目标未完成，仍尝试继续【开始测评】', 'warn');
                await sleep(1200);
            }

            let startBtn = null, noTopic = false;
            for (let i = 0; i < 25; i++) {
                const dialog = evaluationDialogEl();
                startBtn = dialog && findEvaluationButton(/^开始测评$/, dialog);
                if (startBtn) break;
                const msg = evaluationMessageText();
                if (isNoTopicMessage(msg)) { noTopic = true; closeEvaluationMessageBox(); break; }
                if (isEvaluationBlockedMessage(msg)) {
                    log(`平台提示：${msg.slice(0, 160)}`, 'warn');
                    closeEvaluationMessageBox();
                    return { status: 'aborted', reason: 'blocked' };
                }
                await sleep(400);
            }
            if (!startBtn && !noTopic) {
                log('未找到【开始测评】按钮，已停止本次测评', 'warn');
                return { status: 'aborted', reason: 'no-start-button' };
            }
            if (!noTopic) {
                await sleep(600 + Math.floor(Math.random() * 900));
                fireHumanClick(startBtn);
                log('已点击【开始测评】，等待题目加载…', 'info');

                // 等待答题界面或报告出现
                for (let i = 0; i < 50; i++) {
                    if (isEvaluationReportVisible()) { ready = true; break; }
                    const q = readEvaluationQuestionDom();
                    if ((q && q.items && q.items.length) || findEvaluationButton(/^显示报告$/)) { ready = true; break; }
                    const msg = evaluationMessageText();
                    if (isNoTopicMessage(msg)) { noTopic = true; closeEvaluationMessageBox(); break; }
                    if (isEvaluationBlockedMessage(msg)) {
                        log(`平台提示：${msg.slice(0, 160)}`, 'warn');
                        closeEvaluationMessageBox();
                        return { status: 'aborted', reason: 'blocked' };
                    }
                    await sleep(700);
                }
            }
            if (noTopic) {
                log(attempt < 2
                    ? `平台提示【暂无可测评题目】，将重新选择学习目标刷新题池后重试（第 ${attempt + 1}/3 次）`
                    : '重新选择学习目标后仍【暂无可测评题目】', 'warn');
                continue;
            }
            if (!ready) {
                log('测评题目界面未加载：' + describePracticeDom(), 'warn');
                return { status: 'aborted', reason: 'questions-not-ready' };
            }
        }
        if (!ready) {
            return { status: 'aborted', reason: 'no-topic-after-retry' };
        }

        const result = await autoAnswerEvaluationDom();
        if (!result || !result.completed) {
            return { status: 'aborted', reason: evalStopRequested() ? 'manual-stop' : 'incomplete', result };
        }
        // 报告可能比答题流程稍晚渲染，先等报告出现再读取熟练度
        for (let i = 0; i < 20 && !isEvaluationReportVisible(); i++) await sleep(500);
        const box = Array.from(document.querySelectorAll('.ReportBox')).find(isVisibleEl);
        const summary = box ? String(box.innerText || '').replace(/\s+/g, ' ').slice(0, 180) : '';
        if (summary) log(`测评报告：${summary}`, 'success');
        let mastery = extractEvaluationMastery();
        for (let i = 0; i < 8 && mastery == null; i++) {
            await sleep(700);
            mastery = extractEvaluationMastery();
        }
        return { status: 'completed', mastery, summary, result };
    }

    function findActionButton(pattern) {
        const roots = [];
        const lib = document.querySelector('.SubjectLib');
        if (lib) roots.push(lib);
        roots.push(document);
        for (const root of roots) {
            const buttons = Array.from(root.querySelectorAll('button, .el-button'));
            const hit = buttons.find(btn => {
                const text = (btn.innerText || '').trim();
                return pattern.test(text) && isVisibleEl(btn) &&
                    !btn.disabled && !btn.classList.contains('is-disabled');
            });
            if (hit) return hit;
        }
        return null;
    }

    function acquireAiLock() {
        const now = Date.now();
        const lock = Storage.get(AI_LOCK_KEY, null);
        if (lock && lock.id && lock.id !== AI_TAB_ID && now - Number(lock.ts || 0) < 20000) {
            return false;
        }
        Storage.set(AI_LOCK_KEY, { id: AI_TAB_ID, ts: now });
        return true;
    }

    function aiLockHeldByOther() {
        const lock = Storage.get(AI_LOCK_KEY, null);
        return !!(lock && lock.id && lock.id !== AI_TAB_ID && Date.now() - Number(lock.ts || 0) < 20000);
    }

    function touchAiLock() {
        Storage.set(AI_LOCK_KEY, { id: AI_TAB_ID, ts: Date.now() });
    }

    function releaseAiLock() {
        const lock = Storage.get(AI_LOCK_KEY, null);
        if (lock && lock.id === AI_TAB_ID) Storage.set(AI_LOCK_KEY, null);
    }

    // 答完最后一题后点击「显示报告/交卷」
    async function openExerciseReport() {
        for (let i = 0; i < 12 && state.config.running; i++) {
            const btn = findActionButton(/^(显示报告|查看报告|交卷)/);
            if (btn) {
                btn.click();
                log('已点击【显示报告/交卷】', 'success');
                await sleep(2500);
                await handleModals();
                return true;
            }
            await sleep(700);
        }
        log('未找到【显示报告/交卷】按钮，请手动确认', 'warn');
        return false;
    }

    // 核心答题循环：读取题目 -> 请求 AI -> 填入答案 -> 平台判题 -> 下一题
    async function autoAnswerLoop(inst) {
        const st = inst.setupState;
        const minDelay = clampNumber(state.config.aiMinDelayMs, 500, 30000, 2500);
        const maxDelay = Math.max(minDelay, clampNumber(state.config.aiMaxDelayMs, 500, 60000, 6000));
        const maxQuestions = clampNumber(state.config.aiMaxQuestions, 1, 200, 50);
        let answered = 0, correct = 0, wrong = 0, aiErrors = 0, emptyWaits = 0;
        let lastQid = '', sameQidCount = 0;
        const submitted = new Set();

        while (state.config.running && answered < maxQuestions) {
            if (aiLockHeldByOther()) {
                log('另一个标签页已接管答题，本页退出以防重复提交', 'warn');
                return { answered, correct, wrong, completed: false };
            }
            touchAiLock();
            if (!document.querySelector('.SubjectLib')) {
                log('答题界面已关闭，结束本小节', 'warn');
                return { answered, correct, wrong, completed: false };
            }
            if (st.isSubjectOver) break;

            const question = st.Question;
            if (!question || question.id === undefined || question.id === null) {
                emptyWaits++;
                if (emptyWaits > 30) {
                    log('等待题目加载超时，已停止本小节（请确认已点击开始练习）', 'warn');
                    return { answered, correct, wrong, completed: false };
                }
                await sleep(800);
                continue;
            }
            emptyWaits = 0;

            const qid = String(question.id);
            if (qid === lastQid) {
                sameQidCount++;
            } else {
                lastQid = qid;
                sameQidCount = 0;
            }
            if (sameQidCount > 90) {
                log(`题目【${qid}】停留超过 90 秒，已停止本小节以防死循环`, 'error');
                return { answered, correct, wrong, completed: false };
            }

            // 已提交过该题：等待平台自动跳题；答错则手动点「下一题」
            if (submitted.has(String(question.id))) {
                if (st.isAnswered && !st.isCorrect) {
                    const nextBtn = findActionButton(/^下一题/);
                    if (nextBtn) {
                        nextBtn.click();
                        submitted.delete(String(question.id));
                        log('已点击【下一题】', 'info');
                        await sleep(1200);
                        continue;
                    }
                }
                await sleep(1000);
                continue;
            }

            const set = buildQuestionSet(question);
            if (!set) {
                log(`题目【${question.id}】缺少题干或选项，无法自动作答，已停止本小节`, 'warn');
                return { answered, correct, wrong, completed: false };
            }

            const kind = detectQuestionKind(set.items[0].qType);
            const extra = set.items.length > 1 ? `，${set.items.length} 个子题` : '';
            log(`AI 正在作答第 ${answered + 1} 题（${kind}${extra}）...`, 'info');
            touchAiLock();

            let answers;
            try {
                answers = await askAiForQuestionSet(set);
                aiErrors = 0;
            } catch (err) {
                aiErrors++;
                log(`AI 作答失败：${err.message}`, 'error');
                if (aiErrors >= 3) {
                    log('连续 3 次 AI 失败，已暂停自动刷课以保护账号', 'error');
                    state.config.running = false;
                    saveConfig();
                    updatePanelUI();
                    return { answered, correct, wrong, completed: false };
                }
                await sleep(3000);
                continue;
            }

            applyAiAnswers(set, answers);
            const preview = answers.map(a => a.map(i => String.fromCharCode(65 + i)).join('')).join(' / ');
            log(`已填入 AI 答案：${preview}`, 'info');

            // 模拟人类思考时间，降低高频提交风控概率
            const thinkMs = minDelay + Math.floor(Math.random() * (maxDelay - minDelay + 1));
            log(`答案已填入，${(thinkMs / 1000).toFixed(1)} 秒后提交`, 'info');
            await sleep(thinkMs);

            if (!state.config.running) return { answered, correct, wrong, completed: false };
            if (!st.Question || st.Question.id !== question.id) continue;
            if (typeof st.inspect !== 'function') {
                log('未找到平台判题方法 inspect()，无法自动提交', 'error');
                return { answered, correct, wrong, completed: false };
            }

            // 平台判题倒计时期间不重复提交
            for (let i = 0; i < 10 && st.isloadTime; i++) await sleep(500);

            submitted.add(String(question.id));
            touchAiLock();
            try {
                st.inspect();
            } catch (err) {
                submitted.delete(String(question.id));
                log(`调用平台判题失败：${err.message}`, 'error');
                return { answered, correct, wrong, completed: false };
            }
            answered++;

            // 等待平台返回判题结果
            let waited = 0;
            while (waited < 25000 && state.config.running && !st.isAnswered) {
                await sleep(500);
                waited += 500;
            }
            if (!st.isAnswered) {
                log('等待判题结果超时，已停止本小节', 'error');
                return { answered, correct, wrong, completed: false };
            }

            if (st.isCorrect) {
                correct++;
                commitAiCache(true);
                log(`第 ${answered} 题：正确`, 'success');
                // 答对后平台 3 秒自动进入下一题，这里等待题目 ID 变化
                for (let i = 0; i < 24 && state.config.running; i++) {
                    if (st.isSubjectOver) break;
                    if (!st.Question || st.Question.id !== question.id) break;
                    await sleep(500);
                }
                submitted.delete(String(question.id));
            } else {
                wrong++;
                commitAiCache(false);
                log(`第 ${answered} 题：错误（累计正确 ${correct} / 错误 ${wrong}）`, 'warn');
                // 答错时平台不会自动跳题，等倒计时结束后点「下一题」
                for (let i = 0; i < 24 && state.config.running; i++) {
                    if (!st.isloadTime) break;
                    await sleep(500);
                }
                const nextBtn = findActionButton(/^下一题/);
                if (nextBtn) {
                    nextBtn.click();
                    log('已点击【下一题】', 'info');
                } else if (typeof st.getNextSubject === 'function') {
                    st.getNextSubject();
                } else {
                    log('未找到【下一题】按钮，已停止本小节', 'error');
                    return { answered, correct, wrong, completed: false };
                }
                submitted.delete(String(question.id));
                await sleep(1000);
            }
        }

        if (st.isSubjectOver) {
            log(`本节练习作答完成：共 ${answered} 题，正确 ${correct}，错误 ${wrong}`, 'success');
            if (state.config.aiAutoSubmit) await openExerciseReport();
            return { answered, correct, wrong, completed: true };
        }
        if (answered >= maxQuestions) {
            log(`已达到单次答题上限 ${maxQuestions} 题，自动刷课已暂停，重新开始可继续`, 'warn');
            state.config.running = false;
            saveConfig();
            updatePanelUI();
        }
        return { answered, correct, wrong, completed: false };
    }

    async function autoAnswerExercise(inst) {
        if (!acquireAiLock()) {
            log('检测到另一个标签页正在答题，已跳过，避免重复提交', 'warn');
            return null;
        }
        try {
            return await autoAnswerLoop(inst);
        } finally {
            releaseAiLock();
        }
    }

    // 判断当前是否处于练习报告/成绩页（已完成），避免把「重新练题」误当成开始按钮反复重做
    function isExerciseReportPage() {
        if (document.querySelector('.Achievement, .GaugebBoxcn')) return true;
        const text = (document.body.innerText || '').replace(/\s+/g, '');
        return /知识点变化|错题回顾/.test(text) && /重新练[习题]|再练一次/.test(text);
    }

    // 查找练习入口按钮：开始/继续/重新练题（平台对已尝试过的练习会显示「重新练题」）
    function findExerciseStartButton() {
        const pattern = /^(开始练习|继续练习|重新练[习题]|再练一次|再次练习|开始答题|继续答题|开始测验|继续测验|开始检测|继续检测|提升掌握度|提升熟练度|开始提升|去练习)$/;
        const selectors = ['.SubjectCover button', '.SubjectCover .el-button', '.SubjectLayout button', '.el-button', 'button'];
        for (const sel of selectors) {
            const btn = Array.from(document.querySelectorAll(sel)).find(b => {
                if (!isVisibleEl(b) || b.disabled || b.classList.contains('is-disabled')) return false;
                return pattern.test((b.innerText || '').trim());
            });
            if (btn) return btn;
        }
        return null;
    }

    // 通用可点击入口查找：兼容 button / div / span / a 等非 button 元素
    // 平台部分练习入口（如「提升掌握度」）是 div/span 卡片，只按 button 查找会漏掉
    function findClickableByText(pattern) {
        const roots = [];
        const lib = document.querySelector('.SubjectLib');
        if (lib) roots.push(lib);
        roots.push(document);
        const sel = 'button, .el-button, a, div, span, li';
        for (const root of roots) {
            const els = Array.from(root.querySelectorAll(sel));
            const hits = els.filter(el => {
                if (!isVisibleEl(el) || el.disabled || el.classList.contains('is-disabled')) return false;
                const text = String(el.innerText || '').replace(/\s+/g, '');
                if (!text || text.length > 20) return false;
                if (!pattern.test(text)) return false;
                // 只取最内层命中元素，避免点到外层包裹容器
                return !Array.from(el.children).some(c => pattern.test(String(c.innerText || '').replace(/\s+/g, '')));
            });
            if (hits.length) {
                hits.sort((a, b) => {
                    const rank = (el) => el.tagName === 'BUTTON' ? 0 :
                        /btn|button/i.test(String(el.className || '')) ? 1 :
                        el.tagName === 'A' ? 2 : 3;
                    return rank(a) - rank(b);
                });
                const el = hits[0];
                // 命中的是按钮内部节点时，点击动作交给外层按钮
                return el.closest('button, .el-button, a') || el;
            }
        }
        return null;
    }

    /**
     * 处理自适应练习课件
     */
    async function processExerciseActivity(curName) {
        if (!state.config.autoExercise) {
            log(`已跳过自适应练习（配置为不自动作答）：【${curName}】`, 'warn');
            await sleep(1500);
            return;
        }

        if (!state.config.aiEnabled || !String(state.config.aiApiKey || '').trim()) {
            log(`AI 自动答题未启用或未配置 API Key，已跳过练习：【${curName}】`, 'warn');
            await sleep(1500);
            return;
        }

        log(`正在处理自适应练习（AI 自动答题）：【${curName}】`, 'info');
        const sectionName = readCurrentSectionName();
        state.activeExerciseName = curName;
        state.confirmBlockedCount = 0;
        // 平台判定知识点已掌握 / 练习已做过时点练习只弹提示，不会出现题目：直接标记该活动并返回，避免反复切换页面
        // 注意：上一条练习残留的 toast 可能仍在 DOM 中，只有「既无题目、又无任何练习入口」时才认定被平台拦截
        const hasEntryBtn = () => !!(findExerciseStartButton() ||
            findClickableByText(/^(提升掌握度|提升熟练度|开始练习|继续练习|重新练[习题]|再练一次|开始答题|继续答题|去练习)$/));
        const blockedTip = (readPracticeQuestionDom() || hasEntryBtn()) ? '' : readPracticeBlockedTip();
        if (blockedTip) {
            await handleModals();
            markExerciseDone(sectionName, curName, `平台提示：${blockedTip}`);
            await sleep(800);
            return;
        }
        // 平台已把该练习标记为「练过」：只剩练习记录/掌握度变化，没有作答入口 → 标记完成并顺延
        const alreadyDone = detectExerciseAlreadyDone();
        if (alreadyDone) {
            await handleModals();
            markExerciseDone(sectionName, curName, alreadyDone);
            await sleep(800);
            return;
        }
        snapshotPracticeDOM('begin');
        uploadDiag({ tag: 'scripts', ts: new Date().toLocaleTimeString('zh-CN', { hour12: false }), info: describePageScripts() });

        // 1. 若已处于报告/成绩页（通常显示「重新练题」），点击进入新一轮作答
        let entered = false;
        if (isExerciseReportPage()) {
            const redoBtn = findActionButton(/^(重新练[习题]|再练一次|再次练习)/);
            if (redoBtn) {
                fireHumanClick(redoBtn);
                log('检测到【重新练题】，已点击进入新一轮作答...', 'info');
                entered = true;
                await sleep(2500);
                await handleModals();
            } else {
                log(`练习【${curName}】处于报告页，但未找到【重新练题】按钮`, 'warn');
            }
        }

        // 2. 否则查找开始 / 继续 / 重新练题入口（含 div/span 形式的「提升掌握度」卡片）
        if (!entered) {
            const startBtn = findExerciseStartButton() ||
                findClickableByText(/^(提升掌握度|提升熟练度|开始练习|继续练习|重新练[习题]|再练一次|开始答题|继续答题|去练习)$/);
            if (startBtn) {
                const label = (startBtn.innerText || '').trim() || '练习入口';
                fireHumanClick(startBtn);
                log(`已点击【${label}】，等待题目加载...`, 'info');
                await sleep(2500);
            }
        }

        // 处理可能弹出的“仍要学习/重新练习”确认框
        await handleModals();

        // 3. 等待题目界面出现（纯 DOM 判定：题目 / 报告）；入口仍在时继续点击，
        //    兼容「提升掌握度 → 开始练习」这类需要点两层的入口
        let qdom = null;
        let startClickCount = 0;
        let stickyTip = '';
        let stickyTipHits = 0;
        for (let i = 0; i < 40; i++) {
            qdom = readPracticeQuestionDom();
            if (qdom) break;
            // 等题期间一旦出现「已做过/已掌握/无题」提示，立即标记完成并顺延，不再空等。
            // 没有练习入口时可立即认定；入口仍在时要求同一条提示在点击后重复出现才认定，
            // 避免上一条练习残留的 toast 被误当成当前练习已被平台拦截
            const rawTip = readPracticeBlockedTip();
            if (rawTip && rawTip === stickyTip) stickyTipHits++;
            else { stickyTip = rawTip; stickyTipHits = rawTip ? 1 : 0; }
            if (rawTip && (!hasEntryBtn() || (startClickCount >= 1 && stickyTipHits >= 2))) {
                await handleModals();
                markExerciseDone(sectionName, curName, `平台提示：${rawTip}`);
                await sleep(800);
                return;
            }
            // 页面渲染成「练过/已完成」的掌握度记录页（无作答入口）→ 标记完成并顺延
            if (i >= 1) {
                const doneNow = detectExerciseAlreadyDone();
                if (doneNow) {
                    await handleModals();
                    markExerciseDone(sectionName, curName, doneNow);
                    await sleep(800);
                    return;
                }
            }
            if (i === 12) {
                // 12 秒既无题目也无入口，继续等待只会在目录与学习页之间来回切换
                const anyEntry = findPracticeButton(/提升掌握度|提升熟练度|开始练习|继续练习|重新练[习题]|再练一次/) ||
                    findExerciseStartButton();
                if (!anyEntry) {
                    log('等待 12 秒仍无题目与练习入口，提前结束本练习', 'warn');
                    break;
                }
            }
            if (i >= 2 && startClickCount < 4) {
                const coverBtn = findPracticeButton(/提升掌握度|提升熟练度|开始练习|继续练习|重新练[习题]|再练一次/) ||
                    findExerciseStartButton() ||
                    findClickableByText(/^(提升掌握度|提升熟练度|开始练习|继续练习|重新练[习题]|再练一次|去练习)$/);
                if (coverBtn) {
                    startClickCount++;
                    const label = (coverBtn.innerText || '').trim() || '练习入口';
                    fireHumanClick(coverBtn);
                    log(`已点击【${label}】，等待题目加载...`, 'info');
                    await sleep(2000);
                    continue;
                }
            }
            if (i === 5 || i === 15 || i === 25) log('练习界面加载中… ' + describePracticeDom(), 'warn');
            await sleep(1000);
        }

        if (!qdom) {
            log('未检测到练习题目界面，本次跳过本练习并顺延。' + describePracticeDom(), 'warn');
            snapshotPracticeDOM('give-up');
            markExerciseFailed(sectionName, curName, '未出现题目/入口');
            await sleep(1200);
            return;
        }

        // 4. 报告页：尝试【重新练题】获取新题；若仍拿到空报告（0/1 0/0），再试一次
        if (qdom.over) {
            for (let attempt = 0; attempt < 2; attempt++) {
                log(`练习处于报告页（第 ${attempt + 1} 次尝试），点击【重新练题/提升掌握度】...`, 'warn');
                const redoBtn = findPracticeButton(/重新练[习题]|再练一次|再次练习|提升掌握度|提升熟练度/) ||
                    findActionButton(/^(重新练[习题]|再练一次|再次练习|提升掌握度|提升熟练度)/) ||
                    findExerciseStartButton() ||
                    findClickableByText(/^(重新练[习题]|再练一次|再次练习|提升掌握度|提升熟练度)$/);
                if (!redoBtn) break;
                fireHumanClick(redoBtn);
                await sleep(2500);
                await handleModals();
                let got = null;
                for (let i = 0; i < 15; i++) {
                    got = readPracticeQuestionDom();
                    if (got && !got.over) break;
                    if (got && got.over) break;  // 又回到报告页，说明 next 返回 -1
                    await sleep(800);
                }
                if (got && !got.over) {
                    qdom = got;
                    break;  // 成功拿到题目，跳出重试
                }
                if (got && got.over) {
                    // 还是报告页（空报告），继续下一次重试
                    continue;
                }
                // 封面态：再点一次封面按钮
                const coverBtn = findPracticeButton(/提升掌握度|提升熟练度|开始练习|继续练习/) || findExerciseStartButton();
                if (coverBtn) {
                    fireHumanClick(coverBtn);
                    await sleep(2000);
                }
            }
            if (qdom.over) {
                log('该练习两次尝试仍直接返回空报告（平台可能已判定无新题），标记完成并顺延', 'warn');
                markExerciseDone(sectionName, curName, '两次均返回空报告');
                await sleep(1200);
                return;
            }
        }

        // 5. 纯 DOM 自动答题主循环
        const result = await autoAnswerPracticeDom();
        await sleep(1200);
        if (result && result.completed) {
            log(`练习已完成并返回目录：提交 ${result.answered} 题，正确 ${result.correct}，错误 ${result.wrong}`, 'success');
        }
        // 仍有未答/未提交题目时不返回目录，避免触发「正在答题中」弹窗丢弃后面的题；
        // 下一轮主循环会重新进入本节继续作答
        try {
            const stillActive = (() => {
                if (state.evaluationRunning || document.querySelector('.CEDialog, .dialogBox')) return false;
                const q = readPracticeQuestionDom();
                return !!(q && !q.over);
            })();
            if (stillActive) {
                log('本节练习仍有未答题目，留在当前页继续作答（不返回目录）', 'info');
                return;
            }
        } catch (e) { /* ignore */ }
        await returnToCourseIndex();
    }

    /**
     * 章节学习页主状态机
     */
    /**
     * 检测并纠正平台 Vue Router 没有落到目标小节的问题。
     * 平台有时会从 sessionStorage 恢复"上次访问章节"，导致目录页明明点了 A 节但实际进了 B 节。
     * 纠正顺序：
     *   1. 目标 = 当前 → 匹配，清 target，正常继续
     *   2. 找"继续学习下一节"按钮 → 点它
     *   3. 目录侧边栏里再点一次目标 row
     *   4. 有 chapterId → Vue Router.push 硬跳
     *   5. 全部失败 → 回目录让主循环重试
     * 返回 true 表示当前小节已就位（或 target 为空跳过），返回 false 表示已经触发了纠正动作，调用方应 return 等导航完成
     */
    async function ensureCorrectChapter() {
        if (!state.targetSectionName) return true; // 没有目标，跳过校验

        const ch = inspectCurrentChapterDOM();
        const actual = ch.curChapter || '';
        const target = state.targetSectionName;

        // 目标和当前匹配（允许平台加了前缀/后缀，核心名称一致即可）
        const norm = s => String(s || '').trim().replace(/\s+/g, '');
        if (norm(actual) === norm(target) || norm(actual).includes(norm(target)) || norm(target).includes(norm(actual))) {
            state.navFailCount.delete(target);
            state.targetSectionName = '';
            state.targetSectionChapterId = '';
            return true;
        }

        log(`⚠ 平台未落到目标小节！目标：【${target}】，实际：【${actual}】，尝试纠正...`, 'warn');

        // 2. 找"继续学习下一节"按钮（平台常见文案）
        const NEXT_BTN_TEXTS = [
            '继续学习下一节', '继续下一节', '下一节学习', '进入下一节',
            '开始本节学习', '开始学习本节', '学习本节', '学习此节'
        ];
        for (const txt of NEXT_BTN_TEXTS) {
            const btns = Array.from(document.querySelectorAll('button, a, [role="button"], .btn, .el-button, span'));
            const btn = btns.find(b => {
                const t = (b.innerText || b.textContent || '').trim();
                return t && t.includes(txt) && isVisibleEl(b);
            });
            if (btn) {
                log(`找到「${txt}」按钮，点击纠正`, 'info');
                fireHumanClick(btn);
                await sleep(2500);
                // 重新检查
                const ch2 = inspectCurrentChapterDOM();
                const actual2 = ch2.curChapter || '';
                if (norm(actual2) === norm(target) || norm(actual2).includes(norm(target)) || norm(target).includes(norm(actual2))) {
                    state.targetSectionName = '';
                    state.targetSectionChapterId = '';
                    log(`纠正成功：【${actual2}】`, 'success');
                    return true;
                }
            }
        }

        // 2.5 学习页左侧章节树里直接点目标小节（最可靠的纠正方式：
        //     平台侧边栏就是 .tree-node-name，点击等价于目录页点小节）
        const treeNodes = Array.from(document.querySelectorAll('.tree-node-name'));
        const treeHit = treeNodes.find(n => {
            const t = norm(n.innerText);
            return t && (t === norm(target) || t.includes(norm(target)));
        });
        if (treeHit) {
            log(`在学习页侧边栏找到目标小节，点击切换：【${treeHit.innerText.trim().slice(0, 30)}】`, 'info');
            fireHumanClick(treeHit);
            await sleep(3000);
            const chTree = inspectCurrentChapterDOM();
            const actualTree = chTree.curChapter || '';
            if (norm(actualTree) === norm(target) || norm(actualTree).includes(norm(target)) || norm(target).includes(norm(actualTree))) {
                state.navFailCount.delete(target);
                state.targetSectionName = '';
                state.targetSectionChapterId = '';
                log(`侧边栏纠正成功：【${actualTree}】`, 'success');
                return true;
            }
        }

        // 3. 有 chapterId → 用 Vue Router 硬跳
        if (state.targetSectionChapterId) {
            const router = getVueRouter();
            const currentUrl = new URL(location.href);
            currentUrl.pathname = '/istp-learning-center/pattern/chapter';
            currentUrl.searchParams.set('chapterId', state.targetSectionChapterId);
            const courseId = new URLSearchParams(location.search).get('courseId');
            if (courseId) currentUrl.searchParams.set('courseId', courseId);
            const targetUrl = currentUrl.toString();
            log(`尝试 Vue Router.push 到 chapterId=${state.targetSectionChapterId}`, 'info');
            try {
                if (router) {
                    router.push(targetUrl);
                } else {
                    location.href = targetUrl;
                }
                await sleep(2500);
                const ch3 = inspectCurrentChapterDOM();
                const actual3 = ch3.curChapter || '';
                if (norm(actual3) === norm(target) || norm(actual3).includes(norm(target)) || norm(target).includes(norm(actual3))) {
                    state.navFailCount.delete(target);
                    state.targetSectionName = '';
                    state.targetSectionChapterId = '';
                    log(`Router 硬跳纠正成功：【${actual3}】`, 'success');
                    return true;
                }
            } catch (e) { /* ignore */ }
        }

        // 4. 兜底：回目录让主循环重新点目标 row；连续 3 次仍进不去就本次运行跳过该节，
        //    否则会「回目录 → 又点到同一节 → 又落错 → 再回目录」无限循环
        const navFails = (state.navFailCount.get(target) || 0) + 1;
        state.navFailCount.set(target, navFails);
        if (navFails >= 3) {
            state.sessionSkipSections.add(target);
            state.navFailCount.delete(target);
            log(`小节【${target}】连续 ${navFails} 次未落到目标位置，本次运行跳过，改刷下一节`, 'warn');
        }
        log('所有纠正方式都未能落到目标小节，返回目录重试', 'warn');
        await returnToCourseIndex();
        return false;
    }

    async function handleChapterStudyPage() {
        if (!state.config.running) return;

        // 平台自动跳回上次学习小节（我们没有指定目标小节）→ 直接回目录，不做任何学习判断，
        // 否则会把「上次学习的小节」当成待学对象，反复进出已完成小节空跑
        if (!state.targetSectionName && Date.now() < state.indexPinUntil) {
            log('平台自动跳回了上次学习的小节，正在返回课程目录...', 'info');
            await returnToCourseIndex();
            return;
        }

        // 等待加载遮罩消失
        const loadingEl = document.querySelector('.isLoading');
        if (loadingEl && loadingEl.offsetWidth > 0) {
            await sleep(1000);
            return;
        }

        // 检测并纠正 Vue Router 未落到目标小节的情况
        const ok = await ensureCorrectChapter();
        if (!ok) return;

        const ch = inspectCurrentChapterDOM();
        if (!ch.curChapter && ch.totalItems === 0) {
            state.stuckCounter++;
            if (state.stuckCounter > 8) {
                state.stuckCounter = 0;
                await returnToCourseIndex();
            }
            return;
        }
        state.stuckCounter = 0;

        // 切换小节时清空本轮已访问集合
        const chapterKey = `${location.search}__${ch.curChapter}`;
        if (chapterKey !== state.currentChapterKey) {
            state.currentChapterKey = chapterKey;
            state.processedInVisit.clear();
        }

        // 1. 检查跳过名单与单元检测
        if (ch.curChapter && state.handledSections.has(sectionDoneKey(ch.curChapter))) {
            log(`小节【${ch.curChapter}】剩余练习平台已判定掌握，本地标记为已处理，返回目录...`, 'info');
            await returnToCourseIndex();
            return;
        }
        if (ch.curChapter && state.skipSections.has(ch.curChapter)) {
            log(`小节【${ch.curChapter}】在永久跳过列表中，返回目录...`, 'warn');
            await returnToCourseIndex();
            return;
        }
        if (ch.curChapter && state.sessionSkipSections.has(ch.curChapter)) {
            log(`小节【${ch.curChapter}】本次运行已跳过，返回目录...`, 'info');
            await returnToCourseIndex();
            return;
        }
        if (state.config.skipExams && ch.isExam) {
            // 单元检测仅按当前设置跳过，不写入永久跳过列表，关闭该设置后可正常进入
            log(`检测到单元检测【${ch.curChapter || ch.curName}】，按设置自动跳过...`, 'warn');
            await returnToCourseIndex();
            return;
        }
        if (ch.totalItems === 0) {
            // 活动列表可能尚未渲染：允许重试几次，仍为空则本次运行跳过，避免目录↔小节反复横跳
            const key = `__nolist__${ch.curChapter || location.search}`;
            const n = (state.platformDoneVisits.get(key) || 0) + 1;
            state.platformDoneVisits.set(key, n);
            log('当前小节无活动列表，返回课程目录...', 'warn');
            if (n >= 3 && ch.curChapter) {
                state.sessionSkipSections.add(ch.curChapter);
                log(`小节【${ch.curChapter}】连续 ${n} 次未加载出活动列表，本次运行跳过`, 'warn');
            }
            await returnToCourseIndex();
            return;
        }

        // 2. 检查是否全部课件均已完成
        const allFinished = ch.items.every(it => it.isFinish);
        if (allFinished) {
            if (ch.curChapter && state.sectionFailCount.delete(ch.curChapter)) saveSectionFailCounts();
            const doneCount = ch.items.filter(it => it.isDone).length;
            const weakCount = ch.items.filter(it => it.isSessionFailed).length;
            const platformDone = ch.items.filter(it => it.isPlatformDone).length;
            if (ch.curChapter && doneCount > 0 && weakCount === 0) {
                state.handledSections.add(sectionDoneKey(ch.curChapter));
                saveHandledSections();
                log(`小节【${ch.curChapter}】${doneCount} 个练习平台已判定无需再做，本地标记整节为已处理，避免反复进入`, 'warn');
            } else if (ch.curChapter && weakCount > 0) {
                // 有练习本次进不去：不落盘标记整节完成，只在本次运行跳过，刷新后会重试
                state.sessionSkipSections.add(ch.curChapter);
                log(`小节【${ch.curChapter}】有 ${weakCount} 个练习本次无法进入，暂不标记整节完成，刷新后会重试`, 'warn');
            } else if (ch.curChapter && platformDone === ch.items.length) {
                // 平台侧活动全部已完成，但目录页状态可能还没同步：允许重进一次，仍如此则本次运行跳过，避免目录↔小节反复横跳
                const n = (state.platformDoneVisits.get(ch.curChapter) || 0) + 1;
                state.platformDoneVisits.set(ch.curChapter, n);
                if (n >= 2) {
                    state.sessionSkipSections.add(ch.curChapter);
                    log(`小节【${ch.curChapter}】平台侧已全部完成但目录状态未同步，本次运行跳过该小节（刷新后会重新校验）`, 'warn');
                } else {
                    log(`小节【${ch.curChapter}】平台侧活动已全部完成，返回目录等待状态同步（第 ${n} 次）`, 'info');
                }
            }
            const suffix = (doneCount || weakCount) ? `（平台确认无需再做 ${doneCount} 个，本次暂跳过 ${weakCount} 个）` : '';
            log(`★ 小节【${ch.curChapter}】全部 ${ch.totalItems} 个活动已处理${suffix}！返回目录切换下一节`, 'success');
            await returnToCourseIndex();
            return;
        }

        // 3. 筛选未完成且本轮尚未处理过的课件
        const unfinishedIndices = ch.items.map((it, idx) => (!it.isFinish ? idx : -1)).filter(idx => idx >= 0);
        const candidates = unfinishedIndices.filter(idx => !state.processedInVisit.has(idx));

        if (candidates.length === 0 && unfinishedIndices.length > 0) {
            // 平台活动状态回写有延迟：刚处理完的活动在 DOM 上仍是「未完成」，
            // 直接判失败会把实际已完成的整节误记为失败（累计 3 次还会永久跳过）。
            // 这里先等状态刷新，最多 4 次（约 10 秒），仍未回写才判定失败。
            const waitKey = `__status_wait__${ch.curChapter || location.search}`;
            const waited = (state.platformDoneVisits.get(waitKey) || 0) + 1;
            state.platformDoneVisits.set(waitKey, waited);
            if (waited <= 4) {
                log(`小节【${ch.curChapter}】还有 ${unfinishedIndices.length} 个活动未回写完成状态，等待平台刷新（第 ${waited}/4 次）`, 'info');
                await sleep(2500);
                return;
            }
            state.platformDoneVisits.delete(waitKey);
            log(`小节【${ch.curChapter}】所有未完成课件均已尝试处理，触发防死循环保护`, 'warn');
            log(`小节【${ch.curChapter}】仍未回写完成的课件：${unfinishedIndices.map(i => ch.items[i].name).join('、')}`, 'warn');
            markSectionSkipped(ch.curChapter);
            await returnToCourseIndex();
            return;
        }
        if (candidates.length > 0) state.platformDoneVisits.delete(`__status_wait__${ch.curChapter || location.search}`);

        // 4. 若当前激活的不是目标待刷课件，点击切换过去
        const targetIdx = candidates[0];
        if (ch.activeIdx !== targetIdx) {
            // 切活动前清掉上一条活动可能残留的「已掌握/已做过」toast，避免新活动被误判
            clearAllPracticeBlockedToasts();
            // 练习仍有未完成题目时不要切走，否则平台会弹「正在答题中」并冻结脚本
            const liveQ = readPracticeQuestionDom();
            if (liveQ && !liveQ.over) {
                const waited = (state.liveQWaitCount || 0) + 1;
                state.liveQWaitCount = waited;
                if (waited <= 3) {
                    log('当前练习仍有题目未完成，暂不切换活动，等待作答结束', 'warn');
                    await sleep(2000);
                    return;
                }
                // 连续 6 秒都有题未答：说明上一步「空报告 → 标记完成」判断错了，平台其实已经出题。
                // 此时直接对当前练习作答，否则会永久干等（实测卡死在此分支）
                state.liveQWaitCount = 0;
                log('当前练习持续有题目未完成，改为直接作答当前练习', 'warn');
                await processExerciseActivity(ch.curName);
                state.processedInVisit.add(ch.activeIdx);
                return;
            }
            state.liveQWaitCount = 0;
            const targetItem = ch.items[targetIdx];
            log(`切换至待刷课件 [${targetIdx + 1}/${ch.totalItems}]：【${targetItem.name}】`);
            const clickEl = targetItem.el.querySelector('.activity-name, .activity-content-wrapper') || targetItem.el;
            clickEl.click();
            await sleep(2500);
            return;
        }

        // 5. 执行当前激活课件的自动化处理
        if (ch.isVideo) {
            await processVideoActivity(ch.curName);
        } else if (ch.isPPT) {
            await processPPTActivity(ch.curName);
        } else if (ch.isDoc) {
            await processDocActivity(ch.curName);
        } else if (ch.isExercise) {
            await processExerciseActivity(ch.curName);
        } else if (ch.isOtherInstant) {
            log(`正在同步其他资源完成状态：【${ch.curName}】`);
            triggerVueStudyProcess({ total: 1, end: 1, type: '275004' });
            await sleep(2500);
        } else {
            log(`等待课件内容加载：【${ch.curName}】...`);
            await sleep(2000);
        }

        state.processedInVisit.add(ch.activeIdx);
    }

    /**
     * 页面 D: 职业教育 AI 任务引导页 (/istp-learning-center/pattern/task-detail)
     */
    async function handleTaskDetailPage() {
        if (!state.config.running) return;
        const startBtn = document.querySelector('.cta-fixed .cta-btn, .el-button.cta-btn');
        if (startBtn) {
            const t = (startBtn.innerText || '').trim();
            if (t.includes('已完成')) {
                await returnToCourseIndex();
            } else {
                log(`点击任务详情页按钮：【${t}】`);
                startBtn.click();
                await sleep(2500);
            }
        }
    }

    // =========================================================================
    // 8. 悬浮控制面板 UI 与样式注入
    // =========================================================================
    const PANEL_CSS = `
        #fif-helper-panel {
            position: fixed;
            right: 20px;
            bottom: 24px;
            width: min(460px, calc(100vw - 32px));
            background: #1e293b;
            color: #f8fafc;
            border-radius: 12px;
            box-shadow: 0 12px 32px rgba(0, 0, 0, 0.35);
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
            font-size: 13px;
            z-index: 999999;
            border: 1px solid rgba(255, 255, 255, 0.12);
            overflow: hidden;
            transition: height 0.2s ease, width 0.2s ease;
            user-select: none;
        }
        #fif-helper-panel.minimized {
            width: 190px;
        }
        #fif-helper-panel.minimized .fif-body {
            display: none;
        }
        .fif-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px 14px;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            cursor: move;
            font-weight: 600;
            font-size: 13.5px;
        }
        .fif-header-title {
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .fif-badge {
            font-size: 11px;
            padding: 1px 6px;
            border-radius: 99px;
            background: rgba(255, 255, 255, 0.2);
            font-weight: 500;
        }
        .fif-min-btn {
            background: transparent;
            border: none;
            color: #fff;
            cursor: pointer;
            font-size: 15px;
            line-height: 1;
            padding: 2px 6px;
            border-radius: 4px;
        }
        .fif-min-btn:hover {
            background: rgba(255, 255, 255, 0.15);
        }
        .fif-body {
            padding: 12px 14px;
            display: flex;
            flex-direction: column;
            gap: 10px;
            max-height: calc(100vh - 150px);
            overflow-y: auto;
            scrollbar-width: thin;
            scrollbar-color: #475569 #1e293b;
        }
        .fif-body::-webkit-scrollbar,
        .fif-log-list::-webkit-scrollbar {
            width: 10px;
        }
        .fif-body::-webkit-scrollbar-track,
        .fif-log-list::-webkit-scrollbar-track {
            background: rgba(15, 23, 42, 0.6);
            border-radius: 6px;
        }
        .fif-body::-webkit-scrollbar-thumb,
        .fif-log-list::-webkit-scrollbar-thumb {
            background: #475569;
            border-radius: 6px;
            border: 2px solid transparent;
            background-clip: padding-box;
        }
        .fif-body::-webkit-scrollbar-thumb:hover,
        .fif-log-list::-webkit-scrollbar-thumb:hover {
            background: #64748b;
            background-clip: padding-box;
        }
        .fif-status-box {
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 8px 10px;
            font-size: 12px;
            line-height: 1.45;
            color: #38bdf8;
            word-break: break-all;
        }
        .fif-controls {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 6px 10px;
            font-size: 12px;
            color: #cbd5e1;
        }
        .fif-controls label {
            display: flex;
            align-items: center;
            gap: 5px;
            cursor: pointer;
        }
        .fif-rate-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            font-size: 12px;
            color: #cbd5e1;
        }
        .fif-rate-select {
            background: #0f172a;
            color: #f8fafc;
            border: 1px solid #475569;
            border-radius: 5px;
            padding: 2px 6px;
            font-size: 12px;
        }
        .fif-btn-row {
            display: flex;
            gap: 8px;
        }
        .fif-btn {
            flex: 1;
            padding: 7px 10px;
            border: none;
            border-radius: 7px;
            font-weight: 600;
            font-size: 12.5px;
            cursor: pointer;
            transition: opacity 0.15s;
        }
        .fif-btn:hover {
            opacity: 0.92;
        }
        .fif-btn-primary {
            background: #10b981;
            color: #fff;
        }
        .fif-btn-Primary-stop {
            background: #ef4444;
            color: #fff;
        }
        .fif-btn-secondary {
            background: #334155;
            color: #e2e8f0;
            font-size: 11.5px;
            padding: 5px 8px;
        }
        .fif-ai-box {
            border: 1px solid #334155;
            border-radius: 8px;
            padding: 8px 9px;
            display: flex;
            flex-direction: column;
            gap: 6px;
            background: #172033;
        }
        .fif-ai-title {
            display: flex;
            align-items: center;
            justify-content: space-between;
            font-size: 12px;
            font-weight: 600;
            color: #93c5fd;
        }
        .fif-ai-box label {
            display: flex;
            align-items: center;
            gap: 5px;
            font-size: 12px;
            color: #cbd5e1;
            cursor: pointer;
        }
        .fif-ai-box input[type="text"],
        .fif-ai-box input[type="password"],
        .fif-ai-box input[type="number"],
        .fif-ai-box textarea {
            width: 100%;
            box-sizing: border-box;
            background: #0f172a;
            color: #f8fafc;
            border: 1px solid #475569;
            border-radius: 5px;
            padding: 5px 7px;
            font-size: 11.5px;
            font-family: inherit;
            user-select: text;
        }
        .fif-ai-box textarea {
            resize: vertical;
            min-height: 58px;
            line-height: 1.45;
        }
        .fif-ai-row {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 11.5px;
            color: #94a3b8;
        }
        .fif-ai-row input[type="number"] {
            width: 62px;
            flex: none;
        }
        .fif-ai-row .fif-ai-grow {
            flex: 1;
            min-width: 0;
        }
        .fif-ai-test {
            background: #1d4ed8;
            color: #fff;
            border: none;
            border-radius: 5px;
            padding: 5px 8px;
            font-size: 11.5px;
            font-weight: 600;
            cursor: pointer;
        }
        .fif-ai-test:hover {
            opacity: 0.9;
        }
        .fif-risk-box {
            display: flex;
            flex-direction: column;
            gap: 6px;
            background: rgba(127, 29, 29, .28);
            border: 1px solid #b91c1c;
            border-radius: 7px;
            padding: 8px 9px;
        }
        .fif-risk-title {
            font-size: 12px;
            font-weight: 700;
            color: #fca5a5;
        }
        .fif-risk-box label {
            display: flex;
            align-items: center;
            gap: 6px;
            font-size: 12px;
            color: #fecaca;
            cursor: pointer;
        }
        .fif-risk-desc {
            font-size: 11px;
            line-height: 1.5;
            color: #fca5a5;
            word-break: break-all;
        }
        .fif-eval-opts {
            margin-top: 4px;
            font-size: 11px;
            color: #fca5a5;
        }
        .fif-eval-opts label {
            display: flex;
            align-items: center;
            gap: 4px;
            cursor: pointer;
        }
        .fif-eval-hint {
            font-size: 11px;
            line-height: 1.45;
            min-height: 15px;
            color: #94a3b8;
            word-break: break-all;
        }
        .fif-eval-hint.fif-invalid {
            color: #f87171;
            font-weight: 600;
        }
        #fif-cfg-eval-target.fif-input-invalid {
            border-color: #f87171 !important;
            box-shadow: 0 0 0 1px rgba(248, 113, 113, .35);
        }
        .fif-log-list {
            width: 100%;
            box-sizing: border-box;
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 6px;
            height: clamp(240px, 40vh, 420px);
            min-height: 140px;
            resize: vertical;
            overflow-y: auto;
            scrollbar-width: thin;
            scrollbar-color: #475569 #0f172a;
            padding: 6px 8px;
            font-family: Consolas, Monaco, monospace;
            font-size: 11.5px;
            line-height: 1.5;
            color: #94a3b8;
            display: flex;
            flex-direction: column;
            gap: 3px;
        }
        .fif-log-item {
            display: flex;
            align-items: flex-start;
            gap: 4px;
            word-break: break-word;
        }
        .fif-log-item > span {
            flex: 0 0 auto;
            color: #64748b;
        }
        .fif-log-item.success { color: #34d399; }
        .fif-log-item.warn { color: #fbbf24; }
        .fif-helper-card-btn {
            margin-top: 8px;
            width: 100%;
            padding: 6px 10px;
            border: none;
            border-radius: 6px;
            background: linear-gradient(135deg, #2563eb, #1d4ed8);
            color: #fff;
            font-size: 12.5px;
            font-weight: 600;
            cursor: pointer;
            box-shadow: 0 2px 6px rgba(37, 99, 235, 0.35);
            transition: transform 0.12s ease;
        }
        .fif-helper-card-btn:hover {
            transform: translateY(-1px);
            background: linear-gradient(135deg, #1d4ed8, #1e40af);
        }
    `;

    function injectStyles() {
        if (typeof GM_addStyle === 'function') {
            GM_addStyle(PANEL_CSS);
        } else {
            const style = document.createElement('style');
            style.textContent = PANEL_CSS;
            document.head.appendChild(style);
        }
    }

    function createPanel() {
        if (document.getElementById('fif-helper-panel')) return;
        injectStyles();

        const panel = document.createElement('div');
        panel.id = 'fif-helper-panel';
        if (state.config.minimized) panel.classList.add('minimized');

        panel.innerHTML = `
            <div class="fif-header" id="fif-drag-handle">
                <div class="fif-header-title">
                    <span>⚡ FIF全自动刷课助手</span>
                    <span class="fif-badge" id="fif-mode-badge">就绪</span>
                </div>
                <button class="fif-min-btn" id="fif-min-toggle" title="折叠/展开面板">${state.config.minimized ? '＋' : '－'}</button>
            </div>
            <div class="fif-body">
                <div class="fif-status-box" id="fif-status-text">${state.statusText}</div>
                <div class="fif-controls">
                    <label title="遇到自适应练习自动开始，调用 AI 逐题作答并交卷"><input type="checkbox" id="fif-cfg-exercise" ${state.config.autoExercise ? 'checked' : ''}> 自动完成练习</label>
                    <label title="自动跳过所有单元检测与期中期末测验"><input type="checkbox" id="fif-cfg-skip-exam" ${state.config.skipExams ? 'checked' : ''}> 跳过单元检测</label>
                    <label title="仅自动跳转课程/展开课程列表，不伪造学习数据，本身不属于高风险；但会让其他功能在所有课程上无人值守连续执行，放大风控暴露"><input type="checkbox" id="fif-cfg-next-course" ${state.config.autoNextCourse ? 'checked' : ''}> 多课程连刷</label>
                </div>
                <div class="fif-risk-box">
                    <div class="fif-risk-title">⚠ 高风险功能（可能被平台风控识别，谨慎开启）</div>
                    <label title="不实际观看，直接调用 Vue 底层进度接口把视频/文档/PPT 上报为 100% 完成：进度与学习时长严重不符，最容易被平台风控或审计识别，可能导致学习记录作废、账号被限制"><input type="checkbox" id="fif-cfg-fast" ${state.config.fastMode ? 'checked' : ''}> 极速秒刷模式</label>
                    <div class="fif-risk-desc">风险说明：跳过真实观看直接上报 100% 进度，属于最明显的异常学习行为。担心风险时请关闭本项，改用「视频常规播放倍速」。</div>
                </div>
                <div class="fif-rate-row">
                    <span>视频常规播放倍速：</span>
                    <select class="fif-rate-select" id="fif-cfg-rate">
                        <option value="1">1.0x</option>
                        <option value="2">2.0x</option>
                        <option value="4">4.0x</option>
                        <option value="8">8.0x</option>
                        <option value="16">16.0x</option>
                    </select>
                </div>
                <div class="fif-ai-box">
                    <div class="fif-ai-title">
                        <span>AI 自动答题（DeepSeek）</span>
                        <label><input type="checkbox" id="fif-cfg-ai" ${state.config.aiEnabled ? 'checked' : ''}> 启用</label>
                    </div>
                    <input type="password" id="fif-cfg-ai-key" placeholder="DeepSeek API Key" autocomplete="off">
                    <div class="fif-ai-row">
                        <input type="text" id="fif-cfg-ai-model" class="fif-ai-grow" placeholder="模型名称">
                        <button class="fif-ai-test" id="fif-test-ai">测试 API</button>
                    </div>
                    <input type="text" id="fif-cfg-ai-base" placeholder="接口地址（OpenAI 兼容）">
                    <textarea id="fif-cfg-ai-prompt" placeholder="答题提示词"></textarea>
                    <div class="fif-ai-row">
                        <span>思考延迟</span>
                        <input type="number" id="fif-cfg-ai-min" min="500" step="500">
                        <span>~</span>
                        <input type="number" id="fif-cfg-ai-max" min="500" step="500">
                        <span>ms / 上限</span>
                        <input type="number" id="fif-cfg-ai-maxq" min="1" max="200" step="1">
                    </div>
                </div>
                <div class="fif-btn-row">
                    <button class="fif-btn ${state.config.running ? 'fif-btn-Primary-stop' : 'fif-btn-primary'}" id="fif-toggle-run">
                        ${state.config.running ? '⏸ 暂停自动刷课' : '▶ 开始全自动刷课'}
                    </button>
                </div>
                <div class="fif-btn-row">
                    <button class="fif-btn fif-btn-secondary" id="fif-clear-skip" title="清空已标记跳过的异常小节缓存">清空跳过缓存 (<span id="fif-skip-count">${state.skipSections.size}</span>)</button>
                    <button class="fif-btn fif-btn-secondary" id="fif-go-home" title="返回学习空间课程列表首页">返回学习空间</button>
                </div>
                <div class="fif-ai-box" id="fif-eval-box">
                    <div class="fif-ai-title">
                        <span>熟练度测评</span>
                        <label title="开启后：本次测评熟练度低于目标值时自动开始新一轮完整测评；达到目标、手动停止或出现异常时立即终止"><input type="checkbox" id="fif-cfg-eval-loop" ${state.config.evalLoop ? 'checked' : ''}> 循环测试</label>
                    </div>
                    <input type="text" id="fif-cfg-eval-target" inputmode="numeric" placeholder="目标熟练度指标（选填，0-100 整数）" autocomplete="off">
                    <div class="fif-eval-opts">
                        <label title="高风险：忽略目标熟练度，每轮测评结束后自动开始下一轮，直到手动点击「停止测评」；遇到平台拦截（测评间隔过短/已掌握）或连续失败会自动终止"><input type="checkbox" id="fif-cfg-eval-infinite" ${state.config.evalInfinite ? 'checked' : ''}> 无限循环刷熟练度</label>
                    </div>
                    <div class="fif-eval-hint" id="fif-eval-target-hint"></div>
                </div>
                <div class="fif-btn-row">
                    <button class="fif-btn fif-btn-secondary" id="fif-run-eval" title="在课程首页点击「测评评估掌握度」，自动完成测评答题">🎯 测评评估掌握度（自动答题）</button>
                </div>
                <div class="fif-eval-opts">
                    <label title="一键达标「阶段 A」要求每门课章节进度达到的百分比，默认 95；补扫后仍未达标会标记章节完成并继续">章节完成阈值(%) <input type="number" id="fif-cfg-chapter-threshold" min="0" max="100" value="${state.config.fullAutoChapterThreshold}" style="width:60px"></label>
                </div>
                <div class="fif-eval-hint" id="fif-full-auto-phase" style="margin-bottom:6px"></div>
                <div class="fif-btn-row">
                    <button class="fif-btn fif-btn-primary" id="fif-run-full-mastery" title="两阶段：阶段 A 逐门刷章节到设定阈值（补扫后仍不达则标记完成继续），阶段 B 逐门循环测评到目标掌握度（默认最多 8 轮，开启无限循环或设了目标值则按你的设置），全部完成后自动结束">🚀 一键刷全部课程并达标</button>
                </div>
                <div class="fif-log-list" id="fif-log-box"></div>
            </div>
        `;

        document.body.appendChild(panel);

        // 设置默认倍速下拉框
        const rateSelect = document.getElementById('fif-cfg-rate');
        if (rateSelect) rateSelect.value = String(state.config.videoRate || 4);

        // 填充 AI 配置
        const aiKeyInput = document.getElementById('fif-cfg-ai-key');
        const aiModelInput = document.getElementById('fif-cfg-ai-model');
        const aiBaseInput = document.getElementById('fif-cfg-ai-base');
        const aiPromptInput = document.getElementById('fif-cfg-ai-prompt');
        const aiMinInput = document.getElementById('fif-cfg-ai-min');
        const aiMaxInput = document.getElementById('fif-cfg-ai-max');
        const aiMaxQInput = document.getElementById('fif-cfg-ai-maxq');
        if (aiKeyInput) aiKeyInput.value = state.config.aiApiKey || '';
        if (aiModelInput) aiModelInput.value = state.config.aiModel || '';
        if (aiBaseInput) aiBaseInput.value = state.config.aiBaseUrl || '';
        if (aiPromptInput) aiPromptInput.value = state.config.aiPrompt || DEFAULT_AI_PROMPT;
        if (aiMinInput) aiMinInput.value = state.config.aiMinDelayMs;
        if (aiMaxInput) aiMaxInput.value = state.config.aiMaxDelayMs;
        if (aiMaxQInput) aiMaxQInput.value = state.config.aiMaxQuestions;

        // 绑定事件
        document.getElementById('fif-min-toggle').addEventListener('click', () => {
            state.config.minimized = !state.config.minimized;
            saveConfig();
            panel.classList.toggle('minimized', state.config.minimized);
            document.getElementById('fif-min-toggle').textContent = state.config.minimized ? '＋' : '－';
        });

        document.getElementById('fif-toggle-run').addEventListener('click', () => {
            state.config.running = !state.config.running;
            saveConfig();
            log(state.config.running ? '已启动全自动刷课引擎！' : '已暂停自动刷课。', state.config.running ? 'success' : 'warn');
            updatePanelUI();
        });

        document.getElementById('fif-cfg-fast').addEventListener('change', (e) => {
            state.config.fastMode = e.target.checked;
            saveConfig();
            if (state.config.fastMode) {
                log('【高风险】极速秒刷已开启：不观看直接上报 100% 进度，最易被平台风控识别；担心风险请关闭并改用常规倍速', 'warn');
            }
        });
        document.getElementById('fif-cfg-exercise').addEventListener('change', (e) => {
            state.config.autoExercise = e.target.checked;
            saveConfig();
        });
        document.getElementById('fif-cfg-skip-exam').addEventListener('change', (e) => {
            state.config.skipExams = e.target.checked;
            saveConfig();
        });
        document.getElementById('fif-cfg-next-course').addEventListener('change', (e) => {
            state.config.autoNextCourse = e.target.checked;
            saveConfig();
        });
        document.getElementById('fif-cfg-rate').addEventListener('change', (e) => {
            state.config.videoRate = Number(e.target.value) || 4.0;
            saveConfig();
        });

        // ---- AI 答题配置事件 ----
        document.getElementById('fif-cfg-ai').addEventListener('change', (e) => {
            state.config.aiEnabled = e.target.checked;
            saveConfig();
            log(state.config.aiEnabled ? 'AI 自动答题已启用' : 'AI 自动答题已关闭', 'info');
        });
        aiKeyInput.addEventListener('change', (e) => {
            state.config.aiApiKey = e.target.value.trim();
            saveConfig();
            log('DeepSeek API Key 已保存（仅存于本地油猴存储）', 'success');
        });
        aiModelInput.addEventListener('change', (e) => {
            state.config.aiModel = e.target.value.trim() || 'deepseek-chat';
            saveConfig();
        });
        aiBaseInput.addEventListener('change', (e) => {
            state.config.aiBaseUrl = e.target.value.trim() || 'https://api.deepseek.com/chat/completions';
            saveConfig();
        });
        aiPromptInput.addEventListener('change', (e) => {
            state.config.aiPrompt = e.target.value.trim() || DEFAULT_AI_PROMPT;
            saveConfig();
            log('答题提示词已保存', 'success');
        });
        aiMinInput.addEventListener('change', (e) => {
            state.config.aiMinDelayMs = clampNumber(e.target.value, 500, 30000, 2500);
            e.target.value = state.config.aiMinDelayMs;
            if (state.config.aiMaxDelayMs < state.config.aiMinDelayMs) {
                state.config.aiMaxDelayMs = state.config.aiMinDelayMs;
                aiMaxInput.value = state.config.aiMaxDelayMs;
            }
            saveConfig();
        });
        aiMaxInput.addEventListener('change', (e) => {
            state.config.aiMaxDelayMs = clampNumber(e.target.value, 500, 60000, 6000);
            if (state.config.aiMaxDelayMs < state.config.aiMinDelayMs) {
                state.config.aiMaxDelayMs = state.config.aiMinDelayMs;
            }
            e.target.value = state.config.aiMaxDelayMs;
            saveConfig();
        });
        aiMaxQInput.addEventListener('change', (e) => {
            state.config.aiMaxQuestions = clampNumber(e.target.value, 1, 200, 50);
            e.target.value = state.config.aiMaxQuestions;
            saveConfig();
        });
        document.getElementById('fif-test-ai').addEventListener('click', async (e) => {
            const btn = e.currentTarget;
            btn.disabled = true;
            btn.textContent = '测试中...';
            try {
                await testAiConnection();
            } finally {
                btn.disabled = false;
                btn.textContent = '测试 API';
            }
        });

        document.getElementById('fif-clear-skip').addEventListener('click', () => {
            state.skipSections.clear();
            state.doneExercises.clear();
            state.sessionFailedExercises.clear();
            state.exerciseFailCount.clear();
            state.handledSections.clear();
            state.platformDoneVisits.clear();
            state.sectionFailCount.clear();
            state.completedCourses.clear();
            state.masteryMetCourses.clear();
            state.chapterDoneCourses.clear();
            state.chapterRetriedCourses.clear();
            saveSkipSections();
            saveDoneExercises();
            saveHandledSections();
            saveSectionFailCounts();
            saveCompletedCourses();
            saveMasteryMetCourses();
            saveChapterDoneCourses();
            saveChapterRetriedCourses();
            state.config.fullAutoPhase = 'chapters';
            saveConfig();
            log('已清空异常小节、已掌握练习、已完成课程与一键达标全部记录（含阶段/章节进度）！', 'success');
            updatePanelUI();
        });

        document.getElementById('fif-go-home').addEventListener('click', () => {
            location.href = 'https://icourse.fifedu.com/istp-manage/work/home';
        });

        // ---- 熟练度测评配置事件 ----
        const evalLoopInput = document.getElementById('fif-cfg-eval-loop');
        const evalTargetInput = document.getElementById('fif-cfg-eval-target');
        if (evalTargetInput) evalTargetInput.value = state.config.evalTargetMastery || '';
        updateEvalTargetHint();
        if (evalLoopInput) evalLoopInput.addEventListener('change', (e) => {
            state.config.evalLoop = e.target.checked;
            saveConfig();
            updateEvalTargetHint();
            log(state.config.evalLoop
                ? '循环测试已开启：填写目标熟练度后，未达到目标将自动开始新一轮完整测评；达到目标、手动停止或异常时终止'
                : '循环测试已关闭，恢复为单轮测评', 'info');
        });
        if (evalTargetInput) evalTargetInput.addEventListener('input', (e) => {
            state.config.evalTargetMastery = e.target.value;
            saveConfig();
            updateEvalTargetHint();
        });

        const evalInfiniteInput = document.getElementById('fif-cfg-eval-infinite');
        if (evalInfiniteInput) evalInfiniteInput.addEventListener('change', (e) => {
            state.config.evalInfinite = e.target.checked;
            saveConfig();
            updateEvalTargetHint();
            log(state.config.evalInfinite
                ? '【高风险】无限循环已开启：忽略目标熟练度，每轮结束自动开始下一轮，直到手动点击「停止测评」'
                : '无限循环已关闭', state.config.evalInfinite ? 'warn' : 'info');
        });

        // 章节完成阈值（一键达标阶段 A）
        const chapterThresholdInput = document.getElementById('fif-cfg-chapter-threshold');
        if (chapterThresholdInput) {
            chapterThresholdInput.value = state.config.fullAutoChapterThreshold;
            const applyThreshold = (e) => {
                state.config.fullAutoChapterThreshold = clampChapterThreshold(e.target.value);
                saveConfig();
                updatePanelUI();
            };
            chapterThresholdInput.addEventListener('input', applyThreshold);
            chapterThresholdInput.addEventListener('change', applyThreshold);
        }

        document.getElementById('fif-run-eval').addEventListener('click', async (e) => {
            const btn = e.currentTarget;
            if (state.evaluationRunning) {
                state.evalStopRequested = true;
                state.config.running = false;
                log('已请求停止测评：当前步骤结束后立即终止，不再开始新一轮', 'warn');
                updatePanelUI();
                return;
            }
            btn.disabled = true;
            try {
                await runCourseEvaluation();
            } finally {
                btn.disabled = false;
                updatePanelUI();
            }
        });

        // ---- 私有：一键刷全部课程并完成掌握度指标 ----
        const fullMasteryBtnEl = document.getElementById('fif-run-full-mastery');
        if (fullMasteryBtnEl) fullMasteryBtnEl.addEventListener('click', async (e) => {
            const btn = e.currentTarget;

            // 运行中再次点击 = 停止
            if (isFullAutoMastery() || state.evaluationRunning) {
                state.evalStopRequested = true;
                state.config.fullAutoMastery = false;
                state.config.running = false;
                saveConfig();
                updatePanelUI();
                log('已请求停止【一键达标】：当前步骤结束后立即终止，不再切换下一门课程', 'warn');
                return;
            }

            const parsed = parseEvalTargetInput(state.config.evalTargetMastery);
            if (!parsed.provided || !parsed.valid) {
                log('【一键达标】请先在输入框填写合法的目标掌握度（0-100 整数）', 'warn');
                updateEvalTargetHint();
                return;
            }

            state.config.evalLoop = true;
            state.config.autoNextCourse = true;
            state.config.fullAutoMastery = true;
            state.config.fullAutoPhase = 'chapters';
            state.config.running = true;
            state.evalStopRequested = false;
            // 新一轮允许对之前补扫失败的课程再重试一次；章节完成/已达标记录保留以便断点续跑
            state.chapterRetriedCourses.clear();
            saveChapterRetriedCourses();
            saveConfig();
            updatePanelUI();

            const loopEl = document.getElementById('fif-cfg-eval-loop');
            if (loopEl) loopEl.checked = true;
            const th = clampChapterThreshold(state.config.fullAutoChapterThreshold);
            log(`【一键达标】已启动：两阶段执行 —— 阶段 A 逐门刷章节到 ${th}%（补扫仍不达则标记完成继续）→ 阶段 B 逐门循环测评到目标掌握度 ${parsed.value}%（默认最多 8 轮，开启无限循环则按你的设置）`, 'success');

            btn.disabled = true;
            try {
                const path = location.pathname;
                if (path.includes('/istp-learning-center/pattern/chapter') ||
                    path.includes('/istp-learning-center/pattern/task-detail')) {
                    // 章节页无法直接测评，等脚本正常返回目录后自动接续
                    log('【一键达标】当前在章节学习页，完成本章节后会自动返回目录并继续', 'info');
                    return;
                }
                if (path.includes('/istp-learning-center/index')) {
                    await handleCourseIndexPage();
                } else {
                    await handleWorkHomePage();
                }
            } finally {
                btn.disabled = false;
                updatePanelUI();
            }
        });

        // 拖拽支持
        const handle = document.getElementById('fif-drag-handle');
        let isDragging = false, startX = 0, startY = 0, origRight = 20, origBottom = 24;
        handle.addEventListener('mousedown', (e) => {
            if (e.target.id === 'fif-min-toggle') return;
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            const rect = panel.getBoundingClientRect();
            origRight = window.innerWidth - rect.right;
            origBottom = window.innerHeight - rect.bottom;
        });
        window.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;
            panel.style.right = Math.max(0, origRight - dx) + 'px';
            panel.style.bottom = Math.max(0, origBottom - dy) + 'px';
        });
        window.addEventListener('mouseup', () => { isDragging = false; });

        updatePanelUI();
    }

    function updatePanelUI() {
        const statusEl = document.getElementById('fif-status-text');
        const runBtn = document.getElementById('fif-toggle-run');
        const badgeEl = document.getElementById('fif-mode-badge');
        const skipCountEl = document.getElementById('fif-skip-count');
        const logBox = document.getElementById('fif-log-box');

        if (statusEl) statusEl.textContent = state.statusText;
        if (skipCountEl) skipCountEl.textContent = String(state.skipSections.size);
        if (runBtn) {
            runBtn.className = `fif-btn ${state.config.running ? 'fif-btn-Primary-stop' : 'fif-btn-primary'}`;
            runBtn.textContent = state.config.running ? '⏸ 暂停自动刷课' : '▶ 开始全自动刷课';
        }
        if (badgeEl) {
            const path = location.pathname;
            if (path.includes('/istp-manage/work/home')) badgeEl.textContent = '学习空间';
            else if (path.includes('/istp-learning-center/index')) badgeEl.textContent = '章节目录';
            else if (path.includes('/pattern/chapter')) badgeEl.textContent = '刷课中';
            else badgeEl.textContent = state.config.running ? '运行中' : '待机';
        }
        if (logBox) {
            logBox.innerHTML = state.logs.slice(0, 60).map(item =>
                `<div class="fif-log-item ${item.level}"><span>${item.time}</span> ${item.msg}</div>`
            ).join('');
        }
        const evalBtn = document.getElementById('fif-run-eval');
        if (evalBtn) {
            evalBtn.disabled = false;
            evalBtn.textContent = state.evaluationRunning ? '⏹ 停止测评' : '🎯 测评评估掌握度（自动答题）';
            evalBtn.className = `fif-btn ${state.evaluationRunning ? 'fif-btn-Primary-stop' : 'fif-btn-secondary'}`;
            evalBtn.title = state.evaluationRunning ? '立即停止测评/循环测试' : '在课程首页点击「测评评估掌握度」，自动完成测评答题';
        }
        const fullBtn = document.getElementById('fif-run-full-mastery');
        if (fullBtn) {
            fullBtn.disabled = false;
            fullBtn.textContent = isFullAutoMastery() ? '⏹ 停止一键达标' : '🚀 一键刷全部课程并达标';
            fullBtn.className = `fif-btn ${isFullAutoMastery() ? 'fif-btn-Primary-stop' : 'fif-btn-primary'}`;
            fullBtn.title = isFullAutoMastery()
                ? '立即停止一键任务，不再切换下一门课程'
                : '遍历学习空间全部课程，先完成章节学习，再循环测评直到达到目标掌握度，然后自动切换下一门';
        }
        const phaseEl = document.getElementById('fif-full-auto-phase');
        if (phaseEl) {
            if (isFullAutoMastery()) {
                const th = clampChapterThreshold(state.config.fullAutoChapterThreshold);
                if (state.config.fullAutoPhase === 'mastery') {
                    const parsed = parseEvalTargetInput(state.config.evalTargetMastery);
                    phaseEl.textContent = `当前阶段：阶段 B（循环刷掌握度${state.config.evalInfinite ? '，无限循环' : parsed.valid ? `，目标 ${parsed.value}%` : '，无目标值'}）`;
                    phaseEl.style.color = '#c0392b';
                } else {
                    phaseEl.textContent = `当前阶段：阶段 A（逐门刷章节，阈值 ${th}%）`;
                    phaseEl.style.color = '#2980b9';
                }
            } else {
                phaseEl.textContent = '未运行一键达标';
                phaseEl.style.color = '';
            }
        }
    }

    // 目标熟练度输入框的实时校验提示
    function updateEvalTargetHint() {
        const input = document.getElementById('fif-cfg-eval-target');
        const hint = document.getElementById('fif-eval-target-hint');
        if (!hint) return;
        const raw = input ? input.value : state.config.evalTargetMastery;
        const parsed = parseEvalTargetInput(raw);
        if (input) input.classList.toggle('fif-input-invalid', !parsed.valid);
        hint.classList.toggle('fif-invalid', !parsed.valid);
        if (state.config.evalInfinite) {
            hint.textContent = '无限循环已开启：忽略目标熟练度，持续测评直到手动点击「停止测评」'
                + (parsed.valid ? '' : '（当前目标值无效，但无限循环模式下不使用该值）');
            return;
        }
        if (!parsed.valid) {
            hint.textContent = '⚠ ' + parsed.message;
            return;
        }
        if (!parsed.provided) {
            hint.textContent = '选填：0-100 的整数；不填则只进行单轮测评';
        } else if (state.config.evalLoop) {
            hint.textContent = `目标 ${parsed.value}%，未达到将自动再测一轮`;
        } else {
            hint.textContent = `目标 ${parsed.value}%（开启「循环测试」后生效）`;
        }
    }

    // =========================================================================
    // 9. 主循环调度器与油猴菜单注册
    // =========================================================================
    // 一键模式看门狗：无进展且存在残留弹窗时先清理，清理后仍无进展则强制返回学习空间
    async function runFullAutoWatchdog() {
        if (!isFullAutoMastery() || !state.config.running) {
            state.watchdogPhase = 0;
            return false;
        }
        const now = Date.now();
        const hasModal = visibleModals().length > 0;

        if (state.watchdogPhase === 0 && hasModal && now - state.lastProgressAt > 90000) {
            state.watchdogPhase = 1;
            log('一键模式约 90 秒无进展且检测到残留弹窗，正在清理', 'warn');
            await cleanupEvaluationUi('一键模式长时间无进展');
            state.lastProgressAt = Date.now();
            return true;
        }

        if (state.watchdogPhase === 1) {
            if (!hasModal) {
                state.watchdogPhase = 0;
                return false;
            }
            if (now - state.lastProgressAt > 20000) {
                state.watchdogPhase = 2;
                log('清理后弹窗仍未消失，强制返回学习空间继续下一门课程（当前课程测评中断）', 'warn');
                await sleep(600);
                location.href = 'https://icourse.fifedu.com/istp-manage/work/home';
                return true;
            }
        }
        return false;
    }

    async function mainTick() {
        createPanel();
        ensureEvaluationEntry();
        if (!netTapState) startNetTap('always');
        if (!state.config.running) return;  // ← 提前 return，不做任何自动动作
        if (state.evaluationRunning) return;
        if (await runFullAutoWatchdog()) return;
        if (!window.__fifLibWatch) {
            window.__fifLibWatch = true;
            let lastLibHtml = '';
            setInterval(() => {
                try {
                    const lib = document.querySelector('.SubjectLib');
                    if (!lib) return;
                    const html = lib.innerHTML || '';
                    if (html === lastLibHtml) return;
                    lastLibHtml = html;
                    uploadDiag({
                        tag: 'libchange',
                        ts: new Date().toLocaleTimeString('zh-CN', { hour12: false }),
                        hasReport: !!lib.querySelector('.Report'),
                        text: String(lib.innerText || '').replace(/\s+/g, ' ').slice(0, 800),
                        html: html.replace(/\s+/g, ' ').slice(0, 9000)
                    });
                } catch (e) { /* ignore */ }
            }, 1200);
        }
        if (state.busy) return;
        state.busy = true;

        try {
            if (await handleModals()) {
                state.busy = false;
                return;
            }

            const path = location.pathname;
            if (path.includes('/istp-manage/work/home') || path === '/istp-manage/' || path === '/istp-manage') {
                await handleWorkHomePage();
            } else if (path.includes('/istp-learning-center/index')) {
                await handleCourseIndexPage();
            } else if (path.includes('/istp-learning-center/pattern/chapter')) {
                await handleChapterStudyPage();
            } else if (path.includes('/istp-learning-center/pattern/task-detail')) {
                await handleTaskDetailPage();
            }
        } catch (err) {
            console.error('[FIF助手] 调度循环异常:', err);
        } finally {
            state.busy = false;
            updatePanelUI();
        }
    }

    if (typeof GM_registerMenuCommand === 'function') {
        GM_registerMenuCommand('▶ 启动/暂停自动刷课', () => {
            state.config.running = !state.config.running;
            saveConfig();
            updatePanelUI();
        });
        GM_registerMenuCommand('🗑️ 清空跳过记录缓存', () => {
            state.skipSections.clear();
            state.completedCourses.clear();
            saveSkipSections();
            saveCompletedCourses();
            updatePanelUI();
        });
        GM_registerMenuCommand('🎯 测评评估掌握度自动答题', () => {
            runCourseEvaluation();
        });
    }

    // 预安装 Vue Router 跳转守卫（等 Vue 应用挂载后生效）
    setTimeout(installVueRouterRedirector, 800);

    // 主循环里也反复安装（幂等），确保 Vue app 挂载后守卫一定装上
    const routerInstaller = setInterval(() => {
        installVueRouterRedirector();
        if (state.routerRedirectUntil > 0 && Date.now() > state.routerRedirectUntil) {
            // 窗口期过了就停止安装
            // 但 __vueRouterRedirected 已经是 true，setInterval 继续也只是空转
        }
    }, 2000);
    setTimeout(() => clearInterval(routerInstaller), 2 * 60 * 60 * 1000); // 2 小时后停（守卫内部有窗口期判断，空转成本极低）

    // 启动主定时器
    setInterval(mainTick, 2200);
    setTimeout(mainTick, 800);
})();