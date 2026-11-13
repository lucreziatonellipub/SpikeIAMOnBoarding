let interviewDetail = null;
let interviewDetailSource = null;
let interviewDetailFilter = 'all';
let interviewDetailLanguage = 'original';
let interviewDetailReturnFocus = null;
let interviewDetailRequest = 0;
let interviewDetailPreviousOverflow = null;
const interviewSummaryCache = new Map();

function detailAnswerText(value) {
    if (value == null) return '';
    if (typeof value === 'object' && 'answer' in value) return detailAnswerText(value.answer);
    if (typeof value === 'object' && Object.keys(value).length === 0) return '';
    return typeof value === 'object' ? JSON.stringify(value, null, 2) : String(value);
}

function getInterviewSummary(data, checkpoint) {
    const answers = data.collected_data_original || {};
    const english = data.collected_data_english || {};
    // Historical final results don't contain the original questionnaire catalog.
    // Count their saved answers without inventing a historical question total.
    const catalog = Array.isArray(data.questions) && data.questions.length
        ? [...new Set(data.questions)] : null;
    const questions = catalog || [...new Set([...Object.keys(answers), ...Object.keys(english)])];
    const answered = questions.filter(question => detailAnswerText(answers[question] ?? english[question]).trim()).length;
    const facts = {
        company: data.company || 'Not specified',
        target: data.target_system || 'Not specified',
        type: data.system_type || 'Not specified',
        answered,
        total: catalog ? catalog.length : null,
        progress: checkpoint ? data.progress : 100,
        status: data.status_label || (checkpoint ? 'In progress' : 'Completed'),
    };
    const key = `${checkpoint ? 'checkpoint' : 'completed'}:${data.thread_id ?? data.id}`;
    const fingerprint = JSON.stringify(facts);
    const previous = interviewSummaryCache.get(key);
    if (previous?.fingerprint === fingerprint) return previous.text;
    const coverage = facts.total !== null
        ? `${answered} of ${facts.total} questions have been answered.`
        : `${answered} answers have been saved. The original question total is not available.`;
    const text = `Client ${facts.company} is onboarding ${facts.target}, a target of type ${facts.type}. ${coverage} Completion: ${facts.progress}%. Status: ${facts.status}.`;
    interviewSummaryCache.set(key, {fingerprint, text});
    return text;
}

async function openInterviewDetails(id, checkpoint = false, refresh = false) {
    const request = ++interviewDetailRequest;
    try {
        const response = await fetch(checkpoint ? '/api/interviews/checkpoints/' + encodeURIComponent(id) : '/api/sessions/' + id);
        if (!response.ok || response.redirected) throw new Error('Interview unavailable');
        const data = await response.json();
        if (request !== interviewDetailRequest) return;
        const search = refresh ? document.getElementById('detailSearch')?.value || '' : '';
        if (!refresh) {
            interviewDetailReturnFocus = document.activeElement;
            interviewDetailFilter = 'all';
            interviewDetailLanguage = 'original';
        }
        interviewDetail = data;
        interviewDetailSource = {id, checkpoint};
        document.getElementById('sessionModalTitle').textContent = [data.company, data.target_system, data.system_type].filter(Boolean).join(' · ');
        const answers = data.collected_data_original || {};
        const questions = [...new Set([...(data.questions || []), ...Object.keys(answers), ...Object.keys(data.collected_data_english || {})])];
        const answered = questions.filter(q => detailAnswerText(answers[q] ?? data.collected_data_english?.[q]).trim()).length;
        const progress = checkpoint ? data.progress : 100;
        const meta = (label, value) => value ? `<div><dt>${escapeHtml(label)}</dt><dd>${escapeHtml(String(value))}</dd></div>` : '';
        document.getElementById('sessionModalContent').innerHTML = `
            <div class="detail-layout">
                <aside class="detail-summary">
                    <div class="detail-progress-number">${progress}%</div>
                    <div class="detail-progress-caption">Questionnaire completion</div>
                    <progress max="100" value="${progress}" aria-label="Questionnaire completion"></progress>
                    <div class="detail-progress-caption">${answered} answered · ${questions.length - answered} remaining</div>
                    <dl class="detail-metadata">
                        ${meta('Status', data.status_label || 'Completed')}
                        ${meta('Company', data.company)}${meta('Target system', data.target_system)}
                        ${meta('System type', data.system_type)}${meta('User', data.owner)}
                        ${meta('Created', data.created_at ? new Date(data.created_at).toLocaleString() : '')}
                        ${meta('Documents', (data.document_names || []).join(', '))}
                        ${meta('Additional context', (data.other_infos || []).join('\n'))}
                    </dl>
                </aside>
                <section class="detail-main" aria-label="Interview answers">
                    <section class="detail-conversation-summary" aria-labelledby="conversationSummaryTitle">
                        <h3 id="conversationSummaryTitle">Conversation summary</h3>
                        <p>${escapeHtml(getInterviewSummary(data, checkpoint))}</p>
                    </section>
                    <div class="detail-toolbar">
                        <div class="detail-toolbar-top"><h3>Questionnaire & answers</h3><button class="btn detail-refresh" onclick="refreshInterviewDetails()">↻ Refresh</button></div>
                        <input id="detailSearch" type="search" class="detail-search" placeholder="Search questions and answers…" aria-label="Search questions and answers">
                        <div class="detail-tabs">
                            <button class="detail-tab" data-filter="all" onclick="filterInterviewDetails('all')">All</button>
                            <button class="detail-tab" data-filter="answered" onclick="filterInterviewDetails('answered')">Answered</button>
                            <button class="detail-tab" data-filter="pending" onclick="filterInterviewDetails('pending')">To complete</button>
                            <span id="detailVisibleCount" class="detail-count"></span>
                        </div>
                        ${Object.keys(data.collected_data_english || {}).length ? `<div class="detail-tabs"><button class="detail-tab" data-language="original" onclick="setInterviewDetailLanguage('original')">Original answers</button><button class="detail-tab" data-language="english" onclick="setInterviewDetailLanguage('english')">English translation</button></div>` : ''}
                    </div>
                    <div class="detail-answers" id="detailAnswers"></div>
                </section>
            </div>`;
        document.getElementById('detailSearch').value = search;
        document.getElementById('detailSearch').addEventListener('input', renderInterviewAnswers);
        renderInterviewAnswers();
        if (interviewDetailPreviousOverflow === null) {
            interviewDetailPreviousOverflow = document.body.style.overflow;
            document.body.style.overflow = 'hidden';
        }
        document.getElementById('sessionModal').classList.add('active');
        if (!refresh) document.querySelector('#sessionModal .detail-close').focus();
        else document.querySelector('#sessionModal .detail-refresh').focus();
    } catch (error) {
        showToast('Failed to load interview: ' + error.message, 'error');
    }
}

function renderInterviewAnswers() {
    const data = interviewDetail;
    const original = data.collected_data_original || {};
    const answers = interviewDetailLanguage === 'english' ? data.collected_data_english || {} : original;
    const partial = data.document_partial || {};
    const questions = [...new Set([...(data.questions || []), ...Object.keys(original), ...Object.keys(data.collected_data_english || {})])];
    const query = document.getElementById('detailSearch').value.toLocaleLowerCase().trim();
    let visible = 0;
    const html = questions.map((question, index) => {
        const complete = !!detailAnswerText(original[question] ?? data.collected_data_english?.[question]).trim();
        const status = complete ? 'answered' : partial[question] ? 'partial' : 'pending';
        const answer = detailAnswerText(answers[question] ?? original[question] ?? partial[question]);
        if (interviewDetailFilter === 'answered' && !complete) return '';
        if (interviewDetailFilter === 'pending' && complete) return '';
        if (query && !(question + ' ' + answer).toLocaleLowerCase().includes(query)) return '';
        visible++;
        const current = data.current_question === question;
        return `<article class="detail-answer" data-state="${status}">
            <div class="detail-answer-top"><span>Question ${index + 1}${current ? ' · Current question' : ''}</span><span class="detail-answer-state">${complete ? 'Answered' : status === 'partial' ? 'Partial answer' : 'Awaiting answer'}</span></div>
            <h4>${escapeHtml(question)}</h4>
            <div class="detail-answer-text">${escapeHtml(answer || 'No answer has been saved yet.')}</div>
            ${partial[question]?.missing && !complete ? `<div class="detail-note">Still needed: ${escapeHtml(String(partial[question].missing))}</div>` : ''}
        </article>`;
    }).join('');
    document.getElementById('detailAnswers').innerHTML = html || `<div class="detail-empty">${questions.length ? 'No questions match this filter.' : 'No questionnaire questions have been saved yet.'}</div>`;
    document.getElementById('detailVisibleCount').textContent = `${visible} of ${questions.length} questions`;
    document.querySelectorAll('[data-filter]').forEach(button => button.setAttribute('aria-pressed', button.dataset.filter === interviewDetailFilter));
    document.querySelectorAll('[data-language]').forEach(button => button.setAttribute('aria-pressed', button.dataset.language === interviewDetailLanguage));
}

function filterInterviewDetails(filter) {
    interviewDetailFilter = filter;
    renderInterviewAnswers();
}
function setInterviewDetailLanguage(language) {
    interviewDetailLanguage = language;
    renderInterviewAnswers();
}
async function refreshInterviewDetails() {
    if (interviewDetailSource) await openInterviewDetails(interviewDetailSource.id, interviewDetailSource.checkpoint, true);
}
function closeInterviewDetails() {
    ++interviewDetailRequest;
    document.getElementById('sessionModal').classList.remove('active');
    if (interviewDetailPreviousOverflow !== null) document.body.style.overflow = interviewDetailPreviousOverflow;
    interviewDetailPreviousOverflow = null;
    interviewDetailReturnFocus?.focus();
}
document.addEventListener('keydown', event => {
    if (!document.getElementById('sessionModal').classList.contains('active')) return;
    if (event.key === 'Escape') closeInterviewDetails();
    if (event.key === 'Tab') {
        const controls = [...document.querySelectorAll('#sessionModal button, #sessionModal input')].filter(el => !el.disabled);
        const first = controls[0], last = controls[controls.length - 1];
        if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus(); }
        else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus(); }
    }
});
document.getElementById('sessionModal').addEventListener('click', event => {
    if (event.target.id === 'sessionModal') closeInterviewDetails();
});
