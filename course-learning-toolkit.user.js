// ==UserScript==
// @name         FIF智慧教学平台全自动刷课助手 (icourse.fifedu.com)
// @name:zh-CN   FIF智慧教学平台全自动刷课助手 (icourse.fifedu.com)
// @name:en      FIF Smart Teaching Platform Auto Study Helper
// @namespace    https://github.com/XiaoMaxer/course-learning-toolkit
// @version      1.7.2
// @description  适用于 FIF 智慧教学平台 (icourse.fifedu.com) 全站课程自动学习：支持学习空间首页课程识别、全自动巡航章节、视频自动倍速/秒过（含损坏视频容错）、PPT自动翻页完成、Word/PDF文档自动置底完成、自适应练习 AI 自动答题并交卷、课程「测评评估掌握度」自动刷题、自动跳过单元检测及防死循环保护；修复首次进入测评入口不渲染、测评必须手动选择学习目标才能启动的问题；新增测评「循环测试」开关与目标熟练度指标（选填 0-100 整数，实时校验），未达目标自动再测一轮，达标/手动停止/异常即终止。
// @description:zh-CN 适用于 FIF 智慧教学平台 (icourse.fifedu.com) 全站课程自动学习：支持学习空间首页课程识别、全自动巡航章节、视频自动倍速/秒过（含损坏视频容错）、PPT自动翻页完成、Word/PDF文档自动置底完成、自适应练习 AI 自动答题并交卷、课程「测评评估掌握度」自动刷题、自动跳过单元检测及防死循环保护；修复首次进入测评入口不渲染、测评必须手动选择学习目标才能启动的问题；新增测评「循环测试」开关与目标熟练度指标（选填 0-100 整数，实时校验），未达目标自动再测一轮，达标/手动停止/异常即终止。
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
        evalTargetMastery: ''      // 目标熟练度指标（选填，0-100 的整数）
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
            evalTargetMastery: Storage.get('evalTargetMastery', DEFAULT_CONFIG.evalTargetMastery)
        },
        skipSections: new Set(Storage.get('skipSections', [])),
        completedCourses: new Set(Storage.get('completedCourses', [])),
        currentChapterKey: '',
        processedInVisit: new Set(),
        lastIndexUrl: Storage.get('lastIndexUrl', ''),
        busy: false,
        evaluationRunning: false,
        evalStopRequested: false,
        stuckCounter: 0,
        statusText: '就绪，点击「开始刷课」启动自动化',
        logs: []
    };

    function saveConfig() {
        Object.keys(DEFAULT_CONFIG).forEach(k => Storage.set(k, state.config[k]));
    }

    function saveSkipSections() {
        Storage.set('skipSections', Array.from(state.skipSections));
    }

    function saveCompletedCourses() {
        Storage.set('completedCourses', Array.from(state.completedCourses));
    }

    function log(msg, level = 'info') {
        const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
        const line = `[${time}] ${msg}`;
        console.log(`[FIF助手] ${line}`);
        state.logs.unshift({ time, msg, level });
        if (state.logs.length > 60) state.logs.pop();
        state.statusText = msg;
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
    async function handleModals() {
        const isVisible = (el) => {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            return el.offsetWidth > 0 && el.offsetHeight > 0 &&
                   style.display !== 'none' && style.visibility !== 'hidden';
        };
        const btns = Array.from(document.querySelectorAll(
            '.el-message-box__btns button, .el-dialog button, .elMessageBox button, [class*="dialog"] button'
        )).filter(isVisible);

        const targetBtn = btns.find(b => {
            const t = (b.innerText || '').trim();
            return t === '确定' || t === '确认' || t === '我知道了' || t === '仍要学习' || t === '下一步';
        });

        if (targetBtn) {
            const text = targetBtn.innerText.trim();
            targetBtn.click();
            log(`自动点击弹窗按钮：【${text}】`, 'success');
            await sleep(1200);
            return true;
        }
        return false;
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
        // 拦截 window.open 使其在当前标签页或新标签页平滑跳转并保持运行状态
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
            }, 1500);
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

        // 查找尚未 100% 完成且不在已完成列表中的课程卡片
        for (const card of cards) {
            const name = (card.querySelector('.courseName')?.innerText || '').trim();
            const pctText = (card.querySelector('.SelfProgress .percent')?.innerText || '').trim();
            const pctMatch = pctText.match(/(\d+)/);
            const pct = pctMatch ? parseInt(pctMatch[1], 10) : null;

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

    async function handleCourseIndexPage() {
        // 记录当前课程目录页 URL，便于从章节页精准返回
        if (location.href.includes('/istp-learning-center/index')) {
            state.lastIndexUrl = location.href;
            Storage.set('lastIndexUrl', location.href);
        }

        if (!state.config.running) return;

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

        const progEl = document.querySelector('.study-progress-top-left') ||
                       Array.from(document.querySelectorAll('*')).find(e => e.innerText && e.innerText.includes('学习总进度'));
        const totalProgress = progEl ? progEl.innerText.trim().replace(/\s+/g, ' ') : '';

        const todoList = [];
        rows.forEach((row, idx) => {
            const nameEl = row.querySelector('.tree-node-name');
            const resEl = row.querySelector('.right-content');
            if (!nameEl || !resEl) return;

            const isLeaf = nameEl.classList.contains('no-child') || !row.querySelector('.left-arrow');
            const resText = resEl.innerText.trim();
            const m = resText.match(/(\d+)\s*个(资源|学习)活动/);
            const count = m ? parseInt(m[1], 10) : 0;

            const hasCheck = !!row.querySelector('.icon-isfinished, [class*="isfinish"]');
            const isLocked = !!row.querySelector('.icon-locked, .icon-becomedue, [class*="locked"]');
            const name = nameEl.innerText.split('\n')[0].trim();
            const isExam = isExamTitle(name);

            if (isLeaf && count > 0 && !hasCheck && !isLocked) {
                if (state.config.skipExams && isExam) return;
                if (state.skipSections.has(name)) return;
                todoList.push({ idx, name, count, row });
            }
        });

        if (todoList.length === 0) {
            const courseTitle = (document.querySelector('.course-name-warp .name')?.innerText || '').trim();
            if (courseTitle) {
                state.completedCourses.add(courseTitle);
                saveCompletedCourses();
            }
            log(`🎉 当前课程可学小节已全部刷完！(${totalProgress || '100%'})`, 'success');

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
        const clickTarget = target.row.querySelector('.tree-node-name') || target.row;
        clickTarget.click();
        await sleep(3000);
    }

    // =========================================================================
    // 7. 页面 C: 章节学习详情页 (/istp-learning-center/pattern/chapter)
    // =========================================================================
    async function returnToCourseIndex() {
        state.processedInVisit.clear();
        state.currentChapterKey = '';

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
            await sleep(2500);
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
            const isFinish = !!el.querySelector('.is-finish, .is-finished, [class*="isfinish"]') ||
                             statusText.includes('已完成') || statusText.includes('100%');
            return { i, name, statusText, isActive, isFinish, el };
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
     * 处理 PPT 课件
     */
    async function processPPTActivity(curName) {
        log(`正在处理 PPT 课件：【${curName}】`, 'info');
        await sleep(1500);

        // 获取总页数 (如果存在页码显示)
        const totalEl = document.querySelector('.pptContainer .page .total, .btnWrap .page');
        let totalPages = 30;
        if (totalEl) {
            const m = (totalEl.innerText || '').match(/(\d+)\s*$/);
            if (m) totalPages = Math.max(parseInt(m[1], 10), 2);
        }

        // 1. 向当前窗口派发 PreviewPpt 监听的 postMessage 事件 (先发第 1 页解锁首屏标记，再发末页)
        pageWindow.postMessage({
            currentPageNum: 1,
            totalPageNum: totalPages,
            data: { currentPageNum: 1, totalPageNum: totalPages }
        }, '*');
        await sleep(300);
        pageWindow.postMessage({
            currentPageNum: totalPages,
            totalPageNum: totalPages,
            data: { currentPageNum: totalPages, totalPageNum: totalPages }
        }, '*');

        // 2. 点击页面上的下一页按钮 (若存在)
        const nextBtn = document.querySelector('.pptContainer .toggle .pptBtn:last-child, .btnWrap .arrow-right, .btnWrap .pptBtn:last-child');
        if (nextBtn) {
            for (let i = 0; i < Math.min(totalPages, 25); i++) {
                nextBtn.click();
                await sleep(80);
            }
        }

        // 3. 调用 Vue 进度上报接口保底
        triggerVueStudyProcess({ total: totalPages, end: totalPages, type: '275003' });
        log(`PPT 课件已完成：【${curName}】`, 'success');
        await sleep(2500);
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

        try {
            while (state.config.running && answered < maxQuestions) {
                if (aiLockHeldByOther()) {
                    log('另一个标签页已接管答题，本页退出以防重复提交', 'warn');
                    break;
                }
                touchAiLock();

                // 已全部作答（或恢复作答）时，平台会显示「显示报告」，直接进入报告
                const reportBtn = findPracticeButton(/显示报告|查看报告/);
                if (reportBtn) {
                    fireHumanClick(reportBtn);
                    log('已点击【显示报告】', 'info');
                    await sleep(2000);
                    continue;
                }

                const q = readPracticeQuestionDom();
                if (!q) {
                    idle++;
                    if (idle > 40) {
                        log('等待题目超时，结束本节作答', 'warn');
                        break;
                    }
                    await sleep(900);
                    continue;
                }
                if (q.over) break;
                idle = 0;

                const sig = practiceQuestionSignature(q);
                if (doneSigs.has(sig)) {
                    // 作答完成后：最后一题会显示「显示报告」，中间题可能需点「下一题」
                    const advanceBtn = findPracticeButton(/显示报告|查看报告|下一题/);
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

                const answers = await resolveAnswersForDom(q.items);
                if (answers.some(a => !a || !a.length)) {
                    log('未能取得本题可靠答案，暂停作答以避免错误率升高（请检查 AI 配置）', 'warn');
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
        if (over) log(`本节练习作答完成：提交 ${answered} 题，正确 ${correct}，错误 ${wrong}`, 'success');
        return { answered, correct, wrong, completed: over };
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
        let selectAll = findSelectAll();
        if (!selectAll) {
            log('未找到【全选】按钮，已中止本次测评', 'warn');
            return false;
        }
        if (force && /取消全选/.test(String(selectAll.innerText || ''))) {
            // 已全选状态：先取消全选再重新全选，让学习目标发生真实变更，强制平台刷新可测题池
            log('当前已是全选，先【取消全选】再重新【全选】以刷新题池…', 'info');
            fireHumanClick(selectAll);
            for (let i = 0; i < 20; i++) {
                await sleep(300);
                selectAll = findSelectAll();
                if (selectAll && !/取消全选/.test(String(selectAll.innerText || ''))) break;
            }
            await sleep(400);
            selectAll = findSelectAll();
            if (!selectAll) {
                log('取消全选后未找到【全选】按钮，已中止本次测评', 'warn');
                return false;
            }
        }
        if (!/取消全选/.test(String(selectAll.innerText || ''))) fireHumanClick(selectAll);

        let picked = 0;
        for (let i = 0; i < 25; i++) {
            const t = String((setDlg.querySelector('.set-right .set-content-title') || {}).innerText || '');
            const m = t.match(/已选[：:]\s*(\d+)/);
            picked = m ? parseInt(m[1], 10) : 0;
            if (picked > 0) break;
            await sleep(400);
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
        return {
            enabled: !!state.config.evalLoop,
            target,
            // 只有开关开启且目标已填、合法时才进入循环模式，其余情况与原有单轮测评完全一致
            active: !!state.config.evalLoop && target.provided && target.valid
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

    async function runCourseEvaluation() {
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

        const loopCfg = getEvaluationLoopConfig();
        if (loopCfg.enabled && loopCfg.target.provided && !loopCfg.target.valid) {
            log(`目标熟练度指标无效：${loopCfg.target.message}，请修正后再开始测评`, 'warn');
            updateEvalTargetHint();
            return;
        }
        if (loopCfg.enabled && !loopCfg.target.provided) {
            log('循环测试已开启，但未填写目标熟练度指标：本次按单轮测评执行', 'info');
        }
        if (loopCfg.active) {
            log(`循环测试已开启：目标熟练度 ${loopCfg.target.value}%，未达到将自动开始新一轮完整测评；达到目标、手动停止或出现异常时立即终止`, 'info');
        }

        state.evaluationRunning = true;
        state.evalStopRequested = false;
        const prevRunning = state.config.running;
        state.config.running = true;
        updatePanelUI();

        const maxRounds = 20; // 安全上限，避免异常情况下无限循环
        let roundNo = 0;
        try {
            while (true) {
                roundNo++;
                if (evalStopRequested()) {
                    log('已手动停止测评，终止流程', 'warn');
                    break;
                }
                if (loopCfg.active) log(`—— 开始第 ${roundNo} 轮测评 ——`, 'info');

                let round;
                try {
                    round = await runSingleEvaluationRound();
                } catch (err) {
                    log(`第 ${roundNo} 轮测评出现异常：${err && err.message ? err.message : err}，已终止循环`, 'error');
                    break;
                }

                if (!round || round.status !== 'completed') {
                    if (round && round.reason === 'manual-stop') {
                        log('已手动停止测评，终止流程', 'warn');
                    } else if (loopCfg.active) {
                        log(`第 ${roundNo} 轮测评未完成（${round && round.reason ? round.reason : '未知原因'}），终止循环`, 'warn');
                    }
                    break;
                }

                const mastery = round.mastery;
                if (!loopCfg.active) {
                    // 单轮模式：与原有版本完全一致，看完报告后返回课程首页
                    await closeEvaluationReport();
                    break;
                }

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
                if (evalStopRequested()) {
                    log('已手动停止测评，终止循环', 'warn');
                    break;
                }
                if (roundNo >= maxRounds) {
                    log(`已连续测评 ${maxRounds} 轮仍未达到目标，为避免过度请求已终止循环`, 'warn');
                    await closeEvaluationReport();
                    break;
                }

                log(`本次熟练度 ${mastery}% 低于目标 ${loopCfg.target.value}%，将自动开始第 ${roundNo + 1} 轮完整测评（学习目标校验 → 开始测评 → 自动答题 → 查看报告）`, 'info');
                const againBtn = findReportButton(/^再测一次$/);
                if (!againBtn) {
                    log('未找到【再测一次】按钮，无法继续循环，已终止', 'warn');
                    await closeEvaluationReport();
                    break;
                }
                fireHumanClick(againBtn);
                // 等待平台从报告页切回测评封面，再进入下一轮完整流程
                for (let i = 0; i < 20; i++) {
                    await sleep(500);
                    if (!isEvaluationReportVisible()) break;
                }
                await sleep(800);
            }
        } finally {
            state.config.running = prevRunning;
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
        // 强制重走「更改 → 全选 → 确定」刷新题池后重试（最多 2 次）；其余拦截提示仍直接终止。
        let ready = false;
        for (let attempt = 0; attempt < 2 && !ready; attempt++) {
            if (evaluationAnsweringEl() || isEvaluationReportVisible()) { ready = true; break; }
            if (evalStopRequested()) return { status: 'aborted', reason: 'manual-stop' };
            if (attempt > 0) {
                await prepareEvaluationScope(evaluationDialogEl(), true);
                await sleep(600);
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
                log(attempt === 0
                    ? '平台提示【暂无可测评题目】，将重新选择学习目标刷新题池后重试'
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
        const box = Array.from(document.querySelectorAll('.ReportBox')).find(isVisibleEl);
        const summary = box ? String(box.innerText || '').replace(/\s+/g, ' ').slice(0, 180) : '';
        if (summary) log(`测评报告：${summary}`, 'success');
        const mastery = extractEvaluationMastery();
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

        // 2. 否则查找开始 / 继续 / 重新练题入口
        if (!entered) {
            const startBtn = findExerciseStartButton();
            if (startBtn) {
                const label = (startBtn.innerText || '').trim();
                fireHumanClick(startBtn);
                log(`已点击【${label}】，等待题目加载...`, 'info');
                await sleep(2500);
            }
        }

        // 处理可能弹出的“仍要学习/重新练习”确认框
        await handleModals();

        // 3. 等待题目界面出现（纯 DOM 判定：题目 / 报告）；封面按钮出现时再点一次触发练习
        let qdom = null;
        let retriedStart = false;
        for (let i = 0; i < 30; i++) {
            qdom = readPracticeQuestionDom();
            if (qdom && !qdom.over) break;
            if (!retriedStart && i >= 2) {
                const coverBtn = findPracticeButton(/提升掌握度|提升熟练度|开始练习|继续练习|重新练[习题]|再练一次/) ||
                    findExerciseStartButton();
                if (coverBtn) {
                    retriedStart = true;
                    const label = (coverBtn.innerText || '').trim();
                    fireHumanClick(coverBtn);
                    log(`已点击【${label}】，等待题目加载...`, 'info');
                    await sleep(1500);
                    continue;
                }
            }
            if (i === 5 || i === 15) log('练习界面加载中… ' + describePracticeDom(), 'warn');
            await sleep(1000);
        }

        if (!qdom) {
            log('未检测到练习题目界面，已跳过本练习。' + describePracticeDom(), 'warn');
            snapshotPracticeDOM('give-up');
            await sleep(1200);
            await returnToCourseIndex();
            return;
        }

        // 4. 报告页：尝试【重新练题】获取新题；若仍拿到空报告（0/1 0/0），再试一次
        if (qdom.over) {
            for (let attempt = 0; attempt < 2; attempt++) {
                log(`练习处于报告页（第 ${attempt + 1} 次尝试），点击【重新练题/提升掌握度】...`, 'warn');
                const redoBtn = findPracticeButton(/重新练[习题]|再练一次|再次练习|提升掌握度|提升熟练度/) ||
                    findActionButton(/^(重新练[习题]|再练一次|再次练习|提升掌握度|提升熟练度)/) ||
                    findExerciseStartButton();
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
                log('该练习两次尝试仍直接返回空报告（平台可能已判定无新题），返回目录', 'warn');
                await sleep(1200);
                await returnToCourseIndex();
                return;
            }
        }

        // 5. 纯 DOM 自动答题主循环
        const result = await autoAnswerPracticeDom();
        await sleep(1200);
        if (result && result.completed) {
            log(`练习已完成并返回目录：提交 ${result.answered} 题，正确 ${result.correct}，错误 ${result.wrong}`, 'success');
        }
        await returnToCourseIndex();
    }

    /**
     * 章节学习页主状态机
     */
    async function handleChapterStudyPage() {
        if (!state.config.running) return;

        // 等待加载遮罩消失
        const loadingEl = document.querySelector('.isLoading');
        if (loadingEl && loadingEl.offsetWidth > 0) {
            await sleep(1000);
            return;
        }

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
        if (ch.curChapter && state.skipSections.has(ch.curChapter)) {
            log(`小节【${ch.curChapter}】在异常跳过列表中，返回目录...`, 'warn');
            await returnToCourseIndex();
            return;
        }
        if (state.config.skipExams && ch.isExam) {
            log(`检测到单元检测【${ch.curChapter || ch.curName}】，自动跳过...`, 'warn');
            if (ch.curChapter) {
                state.skipSections.add(ch.curChapter);
                saveSkipSections();
            }
            await returnToCourseIndex();
            return;
        }
        if (ch.totalItems === 0) {
            log('当前小节无活动列表，返回课程目录...', 'warn');
            await returnToCourseIndex();
            return;
        }

        // 2. 检查是否全部课件均已完成
        const allFinished = ch.items.every(it => it.isFinish);
        if (allFinished) {
            log(`★ 小节【${ch.curChapter}】全部 ${ch.totalItems} 个活动已完成！返回目录切换下一节`, 'success');
            await returnToCourseIndex();
            return;
        }

        // 3. 筛选未完成且本轮尚未处理过的课件
        const unfinishedIndices = ch.items.map((it, idx) => (!it.isFinish ? idx : -1)).filter(idx => idx >= 0);
        const candidates = unfinishedIndices.filter(idx => !state.processedInVisit.has(idx));

        if (candidates.length === 0 && unfinishedIndices.length > 0) {
            log(`小节【${ch.curChapter}】所有未完成课件均已尝试处理，触发防死循环保护并跳过本节`, 'warn');
            if (ch.curChapter) {
                state.skipSections.add(ch.curChapter);
                saveSkipSections();
            }
            await returnToCourseIndex();
            return;
        }

        // 4. 若当前激活的不是目标待刷课件，点击切换过去
        const targetIdx = candidates[0];
        if (ch.activeIdx !== targetIdx) {
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
            width: 372px;
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
            background: #0f172a;
            border: 1px solid #334155;
            border-radius: 6px;
            height: 92px;
            overflow-y: auto;
            padding: 6px 8px;
            font-family: Consolas, Monaco, monospace;
            font-size: 11px;
            color: #94a3b8;
            display: flex;
            flex-direction: column;
            gap: 3px;
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
                    <div class="fif-eval-hint" id="fif-eval-target-hint"></div>
                </div>
                <div class="fif-btn-row">
                    <button class="fif-btn fif-btn-secondary" id="fif-run-eval" title="在课程首页点击「测评评估掌握度」，自动完成测评答题">🎯 测评评估掌握度（自动答题）</button>
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
            state.completedCourses.clear();
            saveSkipSections();
            saveCompletedCourses();
            log('已清空异常小节与已完成课程跳过记录！', 'success');
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
            logBox.innerHTML = state.logs.slice(0, 25).map(item =>
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
    }

    // 目标熟练度输入框的实时校验提示
    function updateEvalTargetHint() {
        const input = document.getElementById('fif-cfg-eval-target');
        const hint = document.getElementById('fif-eval-target-hint');
        if (!hint) return;
        const raw = input ? input.value : state.config.evalTargetMastery;
        const parsed = parseEvalTargetInput(raw);
        if (!parsed.valid) {
            hint.textContent = '⚠ ' + parsed.message;
            hint.classList.add('fif-invalid');
            if (input) input.classList.add('fif-input-invalid');
            return;
        }
        hint.classList.remove('fif-invalid');
        if (input) input.classList.remove('fif-input-invalid');
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
    async function mainTick() {
        createPanel();
        ensureEvaluationEntry();
        if (!netTapState) startNetTap('always');
        if (state.evaluationRunning) return;
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

    // 启动主定时器
    setInterval(mainTick, 2200);
    setTimeout(mainTick, 800);
})();