let questions = [];
let currentQuestionIndex = 0;
let currentReviewIndex = 0;

// 當頁面載入時，只載入題目但不顯示
window.onload = function() {
    loadQuestions();
};

// 載入題目
function loadQuestions() {
    fetch('questions.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);
            
            questions = jsonData.map(row => ({
                type: row.type.toLowerCase(),
                question: row.question,
                options: row.type.toLowerCase() === 'choice' ? 
                    [row.option1, row.option2, row.option3, row.option4].filter(Boolean) : null,
                answer: String(row.answer)
            }));
        })
        .catch(error => {
            console.error('讀取題目檔案時發生錯誤:', error);
            document.getElementById('currentQuestion').innerHTML = '載入題目時發生錯誤，請重新整理頁面或聯絡管理員。';
        });
}

// 開始測驗
function startQuiz() {
    document.getElementById('startScreen').style.display = 'none';
    document.getElementById('quizContainer').style.display = 'block';
    currentQuestionIndex = 0;
    resetQuizState();
    displayCurrentQuestion();
    updateNavigationButtons();
}

// 重新開始測驗
function restartQuiz() {
    document.getElementById('result').style.display = 'none';
    document.getElementById('reviewContainer').style.display = 'none';
    document.getElementById('quizContainer').style.display = 'block';
    currentQuestionIndex = 0;
    resetQuizState();
    displayCurrentQuestion();
    updateNavigationButtons();
}

// 重設測驗狀態
function resetQuizState() {
    questions.forEach(q => {
        q.userAnswer = undefined;
    });
    updateProgress();
}

function displayCurrentQuestion() {
    const q = questions[currentQuestionIndex];
    const container = document.getElementById('currentQuestion');
    const questionNumber = document.getElementById('questionNumber');
    
    questionNumber.textContent = `第 ${currentQuestionIndex + 1} 題 / 共 ${questions.length} 題`;
    
    container.innerHTML = '';
    const questionDiv = document.createElement('div');
    questionDiv.className = 'question';

    if (q.type === 'choice') {
        questionDiv.innerHTML = `
            <p>${q.question}</p>
            ${q.options.map((option, i) => `
                <label>
                    <input type="radio" name="q${currentQuestionIndex}" value="${option}">
                    ${option}
                </label>
            `).join('<br>')}
        `;
    } else {
        questionDiv.innerHTML = `
            <p>${q.question}</p>
            <input type="text" name="q${currentQuestionIndex}">
        `;
    }

    container.appendChild(questionDiv);
    
    // 恢復之前的答案（如果有）
    const previousAnswer = q.userAnswer;
    if (previousAnswer) {
        if (q.type === 'choice') {
            const radio = container.querySelector(`input[value="${previousAnswer}"]`);
            if (radio) radio.checked = true;
        } else {
            const input = container.querySelector('input[type="text"]');
            if (input) input.value = previousAnswer;
        }
    }

    updateProgress();
}

function updateNavigationButtons() {
    document.getElementById('prevBtn').disabled = currentQuestionIndex === 0;
    const isLastQuestion = currentQuestionIndex === questions.length - 1;
    document.getElementById('nextBtn').style.display = isLastQuestion ? 'none' : 'block';
    document.getElementById('submitBtn').style.display = isLastQuestion ? 'block' : 'none';
}

function saveCurrentAnswer() {
    const q = questions[currentQuestionIndex];
    if (q.type === 'choice') {
        const selected = document.querySelector(`input[name="q${currentQuestionIndex}"]:checked`);
        q.userAnswer = selected ? selected.value : undefined;
    } else {
        const answer = document.querySelector(`input[name="q${currentQuestionIndex}"]`).value;
        q.userAnswer = answer;
    }
}

function previousQuestion() {
    if (currentQuestionIndex > 0) {
        saveCurrentAnswer();
        currentQuestionIndex--;
        displayCurrentQuestion();
        updateNavigationButtons();
    }
}

function nextQuestion() {
    if (currentQuestionIndex < questions.length - 1) {
        saveCurrentAnswer();
        currentQuestionIndex++;
        displayCurrentQuestion();
        updateNavigationButtons();
    }
}

function submitQuiz() {
    saveCurrentAnswer();
    let score = 0;
    
    questions.forEach((q) => {
        if (q.userAnswer === q.answer) {
            score++;
        }
    });

    const finalScore = (score / questions.length) * 100;
    document.getElementById('score').textContent = finalScore.toFixed(1);
    
    const message = document.getElementById('message');
    if (finalScore >= 90) {
        message.textContent = '你超棒的！';
    } else if (finalScore < 60) {
        message.textContent = '還要再多加油喔！';
    } else {
        message.textContent = '做得不錯！';
    }

    document.getElementById('result').style.display = 'block';
    document.getElementById('quizContainer').style.display = 'none';
}

function updateProgress() {
    const progress = ((currentQuestionIndex + 1) / questions.length) * 100;
    document.getElementById('progress').style.width = `${progress}%`;
    document.getElementById('progressText').textContent = `${Math.round(progress)}%`;
}

function showReview() {
    document.getElementById('result').style.display = 'none';
    document.getElementById('reviewContainer').style.display = 'block';
    currentReviewIndex = 0;
    displayReview();
}

function displayReview() {
    const q = questions[currentReviewIndex];
    const container = document.getElementById('reviewContent');
    
    let content = `
        <div class="question">
            <h3>第 ${currentReviewIndex + 1} 題</h3>
            <p>${q.question}</p>
    `;

    if (q.type === 'choice') {
        content += `<div class="options-review">`;
        q.options.forEach(option => {
            let optionClass = '';
            let statusIcon = '';
            
            if (option === q.answer) {
                optionClass = 'correct-answer';
                statusIcon = '✓ (正確答案)';
            } else if (option === q.userAnswer) {
                optionClass = 'wrong-answer';
                statusIcon = '✗ (你的答案)';
            }

            content += `
                <div class="option ${optionClass}">
                    ${option} ${statusIcon}
                </div>
            `;
        });
        content += `</div>`;
        
        content += `
            <div class="answer-explanation">
                <p>你的答案：${q.userAnswer || '未作答'}</p>
                <p>正確答案：${q.answer}</p>
                <p class="${q.userAnswer === q.answer ? 'correct-answer' : 'wrong-answer'}">
                    ${q.userAnswer === q.answer ? '答對了！' : '答錯了！'}
                </p>
            </div>
        `;
    } else {
        content += `
            <div class="fill-answer">
                <p>你的答案：<span class="${q.userAnswer === q.answer ? 'correct-answer' : 'wrong-answer'}">${q.userAnswer || '未作答'}</span></p>
                <p>正確答案：<span class="correct-answer">${q.answer}</span></p>
            </div>
        `;
    }

    content += `</div>`;
    container.innerHTML = content;
    updateReviewNavigation();
}

function previousReview() {
    if (currentReviewIndex > 0) {
        currentReviewIndex--;
        displayReview();
    }
}

function nextReview() {
    if (currentReviewIndex < questions.length - 1) {
        currentReviewIndex++;
        displayReview();
    }
}

function updateReviewNavigation() {
    document.getElementById('prevReviewBtn').disabled = currentReviewIndex === 0;
    document.getElementById('nextReviewBtn').disabled = currentReviewIndex === questions.length - 1;
}