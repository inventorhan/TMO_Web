document.addEventListener('DOMContentLoaded', () => {
    // HTML 요소 가져오기
    const videoPlayer = document.getElementById('videoPlayer');
    const videoUpload = document.getElementById('videoUpload');
    const workTableBtn = document.getElementById('workTableBtn');
    const analysisBtn = document.getElementById('analysisBtn');
    const manualBtn = document.getElementById('manualBtn');
    const languageSelector = document.getElementById('languageSelector');
    const loadVideoBtn = document.getElementById('loadVideoBtn');
    const timeDisplay = document.getElementById('timeDisplay');
    const videoPathDisplay = document.getElementById('videoPathDisplay');
    
    const unitSelector = document.getElementById('unitSelector');
    const addUnitBtn = document.getElementById('addUnitBtn');
    const editUnitBtn = document.getElementById('editUnitBtn');
    const deleteUnitBtn = document.getElementById('deleteUnitBtn');

    const unitModal = document.getElementById('unitModal');
    const modalTitle = document.getElementById('modalTitle');
    const unitNameInput = document.getElementById('unitNameInput');
    const modalOkBtn = document.getElementById('modalOkBtn');
    const modalCancelBtn = document.getElementById('modalCancelBtn');

    const saveExcelBtn = document.getElementById('saveExcelBtn');
    const loadExcelBtn = document.getElementById('loadExcelBtn');
    const excelUpload = document.getElementById('excelUpload');
    
    const manualModal = document.getElementById('manualModal');
    const closeManualBtn = document.querySelector('.close-manual-btn');
    
    const workTable = document.querySelector('.work-table');
    const analysisView = document.getElementById('analysisView');
    const analysisTableContainer = document.getElementById('analysisTableContainer');
    const analysisChartContainer = document.getElementById('analysisChartContainer');
    const videoSection = document.querySelector('.video-section');
    
    const workTableBody = document.getElementById('workTableBody');
    let currentModalAction = null; // 모달의 'add' 또는 'edit' 상태 저장
    let isVideoFocused = false; // 동영상 제어 단축키 활성화 상태
    

    // --- Unit별 데이터 관리 ---
    const unitDataStore = new Map();

    function initializeUnitData(unitName) {
        if (!unitDataStore.has(unitName)) {
            unitDataStore.set(unitName, {
                videoSrc: null,
                videoName: '선택된 파일 없음',
                tableData: [] // 테이블 행 데이터 배열
            });
        }
    }

    // --- 동영상 관련 기능 ---
    loadVideoBtn.addEventListener('click', () => {
        videoUpload.click(); // 숨겨진 file input을 클릭
    });

    videoUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (file) {
            const fileURL = URL.createObjectURL(file);
            videoPlayer.src = fileURL;
            videoPathDisplay.textContent = file.name;

            // 현재 선택된 Unit의 데이터 저장
            const currentUnit = unitSelector.value;
            unitDataStore.get(currentUnit).videoSrc = fileURL;
            unitDataStore.get(currentUnit).videoName = file.name;
        }
    });

    // --- 다국어 지원 ---
    const translations = {
        'work-table-btn': { ko: '작업 테이블', en: 'Work Table', zh: '工作台' },
        'analysis-btn': { ko: '분석', en: 'Analysis', zh: '分析' },
        'manual-btn': { ko: '설명서', en: 'Manual', zh: '手册' },
        'video-open-btn': { ko: '비디오 열기', en: 'Video Open', zh: '打开视频' },
        'no-file-selected': { ko: '선택된 파일 없음', en: 'No file selected', zh: '未选择文件' },
        'unit-select-label': { ko: 'Unit 선택:', en: 'Select Unit:', zh: '选择单元:' },
        'add-btn': { ko: '추가', en: 'Add', zh: '添加' },
        'edit-btn': { ko: '수정', en: 'Edit', zh: '编辑' },
        'delete-btn': { ko: '삭제', en: 'Delete', zh: '删除' },
        'file-path-label': { ko: '파일 경로:', en: 'File Path:', zh: '文件路径:' },
        'saveExcelBtn': { ko: 'Excel 저장', en: 'Save Excel', zh: '保存Excel' },
        'loadExcelBtn': { ko: 'Excel 불러오기', en: 'Load Excel', zh: '加载Excel' },
        'resetBtn': { ko: 'Reset', en: 'Reset', zh: '重置' },
        'th-unit': { ko: 'Unit', en: 'Unit', zh: '单元' },
        'th-process': { ko: 'Process', en: 'Process', zh: '过程' },
        'th-cycle': { ko: 'Cycle', en: 'Cycle', zh: '周期' },
        'th-work-element': { ko: 'Work Element', en: 'Work Element', zh: '工作要素' },
        'th-unit-motion': { ko: 'Unit Motion', en: 'Unit Motion', zh: '单元动作' },
        'th-start-time': { ko: 'Start Time', en: 'Start Time', zh: '开始时间' },
        'th-end-time': { ko: 'End Time', en: 'End Time', zh: '结束时间' },
        'th-time-interval': { ko: 'Time Interval', en: 'Time Interval', zh: '时间间隔' },
        'th-value-type': { ko: 'Value Type', en: 'Value Type', zh: '价值类型' },
        'modal-title-add': { ko: '새 Unit 추가', en: 'Add New Unit', zh: '添加新单元' },
        'modal-title-edit': { ko: 'Unit 이름 수정', en: 'Edit Unit Name', zh: '编辑单元名称' },
        'modal-placeholder': { ko: 'Unit 이름을 입력하세요', en: 'Enter Unit name', zh: '输入单元名称' },
        'modal-ok': { ko: '확인', en: 'OK', zh: '确认' },
        'modal-cancel': { ko: '취소', en: 'Cancel', zh: '取消' },
        'analysis-title': { ko: 'Unit별 시간 분석', en: 'Time Analysis by Unit', zh: '按单元进行时间分析' },
        'analysis-th-unit': { ko: 'Unit', en: 'Unit', zh: '单元' },
        'analysis-th-va': { ko: 'VA 시간 (초)', en: 'VA Time (s)', zh: 'VA 时间 (秒)' },
        'analysis-th-bva': { ko: 'BVA 시간 (초)', en: 'BVA Time (s)', zh: 'BVA 时间 (秒)' },
        'analysis-th-nva': { ko: 'NVA 시간 (초)', en: 'NVA Time (s)', zh: 'NVA 时间 (秒)' },
        'analysis-th-total': { ko: '총 시간 (초)', en: 'Total Time (s)', zh: '总时间 (秒)' },
        'manual-title': { ko: 'TMO 사용 설명서', en: 'TMO User Manual', zh: 'TMO 用户手册' },
        'manual-video-control': { ko: '동영상 제어', en: 'Video Control', zh: '视频控制' },
        'manual-video-play': { ko: '<b>재생/일시정지:</b> <code>Space Bar</code>를 누르거나, 동영상 하단 컨트롤 바를 이용하세요.', en: '<b>Play/Pause:</b> Press the <code>Space Bar</code> or use the control bar at the bottom of the video.', zh: '<b>播放/暂停:</b> 按 <code>空格键</code> 或使用视频底部的控制栏。' },
        'manual-video-load': { ko: '<b>불러오기:</b> [비디오 열기] 버튼을 클릭하거나, 동영상 파일을 플레이어 영역으로 드래그 앤 드롭하세요.', en: '<b>Load:</b> Click the [Video Open] button or drag and drop a video file onto the player area.', zh: '<b>加载:</b> 单击 [打开视频] 按钮或将视频文件拖放到播放器区域。' },
        'manual-video-shortcuts-title': { ko: '<b>시간 이동 단축키 (동영상 영역 클릭 후 활성화):</b>', en: '<b>Time Seek Shortcuts (active after clicking video area):</b>', zh: '<b>时间跳转快捷键 (单击视频区域后激活):</b>' },
        'manual-shortcut-qw': { ko: '<b>Q/W:</b> -1초 / +1초', en: '<b>Q/W:</b> -1s / +1s', zh: '<b>Q/W:</b> -1秒 / +1秒' },
        'manual-shortcut-as': { ko: '<b>A/S:</b> -0.1초 / +0.1초', en: '<b>A/S:</b> -0.1s / +0.1s', zh: '<b>A/S:</b> -0.1秒 / +0.1秒' },
        'manual-shortcut-zx': { ko: '<b>Z/X:</b> -0.01초 / +0.01초', en: '<b>Z/X:</b> -0.01s / +0.01s', zh: '<b>Z/X:</b> -0.01秒 / +0.01秒' },
        'manual-table-ops': { ko: '테이블 작업', en: 'Table Operations', zh: '表格操作' },
        'manual-table-add': { ko: '<b>행 추가</b> (동영상 더블클릭): 현재 동영상 시간에 행이 추가되고, 자동으로 선택 및 하이라이트됩니다.', en: '<b>Add Row</b> (Double-click video): Adds a row at the current video time. It will be automatically selected and highlighted.', zh: '<b>添加行</b> (双击视频): 在当前视频时间添加一行，该行将被自动选中并高亮显示。' },
        'manual-table-delete': { ko: '<b>행 삭제</b> (<code>Delete</code> 키): 테이블에서 행을 클릭하여 선택한 후, 키보드의 <code>Delete</code> 키를 누르세요.', en: '<b>Delete Row</b> (<code>Delete</code> key): Select a row in the table by clicking it, then press the <code>Delete</code> key.', zh: '<b>删除行</b> (<code>Delete</code> 键): 在表格中单击以选择一行，然后按 <code>Delete</code> 键。' },
        'manual-table-repeat': { ko: '<b>구간 반복 재생</b> (<code>R</code> 키): 행을 선택한 후, <code>R</code> 키를 누르면 해당 행의 Start Time부터 End Time까지 재생됩니다.', en: '<b>Repeat Playback</b> (<code>R</code> key): Select a row, then press the <code>R</code> key to play the section from its Start Time to End Time.', zh: '<b>重复播放</b> (<code>R</code> 键): 选择一行，然后按 <code>R</code> 键可播放从其开始时间到结束时间的部分。' },
        'manual-table-seek': { ko: '<b>동영상 이동</b> (행 클릭): 행을 클릭하면 해당 행의 \'Start Time\'으로 동영상이 즉시 이동합니다.', en: '<b>Seek Video</b> (Row click): Clicking a row instantly moves the video to that row\'s \'Start Time\'.', zh: '<b>视频跳转</b> (单击行): 单击一行可立即将视频移动到该行的“开始时间”。' },
        'manual-interval-control': { ko: 'Time Interval 조절', en: 'Time Interval Adjustment', zh: '时间间隔调整' },
        'manual-interval-adjust': { ko: '입력 상자의 위/아래 버튼 또는 마우스 휠 스크롤로 0.01초 단위 정밀 조절이 가능합니다.', en: 'Precise 0.01-second adjustments can be made with the input box arrows or the mouse wheel.', zh: '可以使用输入框箭头或鼠标滚轮进行精确的0.01秒调整。' },
        'manual-interval-interlock': { ko: '조절 시 인접한 행의 시간이 자동으로 변경되며, 작업 시간이 서로 겹치지 않도록 안전장치(인터록)가 작동합니다.', en: 'Adjusting it automatically changes the time of adjacent rows, and a safety interlock prevents time overlaps.', zh: '调整它会自动更改相邻行的时间，并且安全联锁可防止时间重叠。' },
        'manual-unit-management': { ko: 'Unit 관리', en: 'Unit Management', zh: '单元管理' },
        'manual-unit-concept': { ko: '각 Unit은 독립된 동영상과 테이블 데이터를 가집니다. Unit을 변경하면 작업 내용이 자동으로 저장되고 복원됩니다.', en: 'Each Unit has its own independent video and table data. Switching Units automatically saves and restores your work.', zh: '每个单元都有自己独立的视频和表格数据。切换单元会自动保存和恢复您的工作。' },
        'manual-unit-purpose': { ko: '<b>사용 목적:</b> 여러 작업 구간, 다른 작업자, 또는 개선 전/후 비교 등 독립적인 분석이 필요할 때 유용합니다. 각 Unit은 별도의 작업 공간 역할을 하여 체계적인 데이터 관리를 돕습니다.', en: '<b>Purpose:</b> Useful for independent analysis, such as comparing different work sections, different workers, or before/after improvements. Each Unit acts as a separate workspace, aiding in organized data management.', zh: '<b>用途:</b> 当需要进行独立分析时（例如比较不同的工作区域、不同的工人或改进前后的情况），此功能非常有用。每个单元都充当一个独立的工作区，有助于进行有组织的数据管理。' },
        'manual-data-io': { ko: '데이터 저장/불러오기', en: 'Data Save/Load', zh: '数据保存/加载' },
        'manual-data-excel': { ko: '[Excel 저장/불러오기] 버튼을 사용하여 모든 Unit의 작업 데이터를 .xlsx 파일로 관리할 수 있습니다. (Excel 불러오기 시, 동영상은 수동으로 다시 선택해야 합니다.)', en: 'You can manage all Unit data as an .xlsx file using the [Save/Load Excel] buttons. (When loading from Excel, the video must be re-selected manually.)', zh: '您可以使用 [保存/加载Excel] 按钮将所有单元数据作为.xlsx文件进行管理。（从Excel加载时，必须手动重新选择视频。）' },
        'manual-data-reset': { ko: '[Reset] 버튼으로 모든 작업 내용을 초기화할 수 있습니다.', en: 'You can clear all work data with the [Reset] button.', zh: '您可以使用 [重置] 按钮清除所有工作数据。' },
    };

    function updateLanguage(lang) {
        document.querySelectorAll('[data-lang-key]').forEach(el => {
            const key = el.getAttribute('data-lang-key');
            if (translations[key] && translations[key][lang]) {
                el.innerHTML = translations[key][lang];
            }
        });
        // 플레이스홀더 등 다른 속성도 변경
        unitNameInput.placeholder = translations['modal-placeholder'][lang];
        modalOkBtn.textContent = translations['modal-ok'][lang];
        modalCancelBtn.textContent = translations['modal-cancel'][lang];
    }

    languageSelector.addEventListener('change', (e) => {
        updateLanguage(e.target.value);
    });

    // --- 설명서 모달 기능 ---
    manualBtn.addEventListener('click', () => {
        manualModal.style.display = 'flex';
    });

    closeManualBtn.addEventListener('click', () => {
        manualModal.style.display = 'none';
    });

    // --- 뷰 전환 (작업 테이블 / 분석) ---
    workTableBtn.addEventListener('click', () => {
        analysisView.style.display = 'none';
        workTable.style.display = 'table';
    });

    analysisBtn.addEventListener('click', () => {
        workTable.style.display = 'none';
        analysisView.style.display = 'block';
        renderAnalysisView();
    });

    function renderAnalysisView() {
        const lang = languageSelector.value;
        let tableHtml = '<table class="analysis-table">';
        tableHtml += `
            <thead>
                <tr>
                    <th data-lang-key="analysis-th-unit">${translations['analysis-th-unit'][lang]}</th>
                    <th data-lang-key="analysis-th-va">${translations['analysis-th-va'][lang]}</th>
                    <th data-lang-key="analysis-th-bva">${translations['analysis-th-bva'][lang]}</th>
                    <th data-lang-key="analysis-th-nva">${translations['analysis-th-nva'][lang]}</th>
                    <th data-lang-key="analysis-th-total">${translations['analysis-th-total'][lang]}</th>
                </tr>
            </thead>
            <tbody>
        `;

        analysisChartContainer.innerHTML = ''; // 차트 컨테이너 초기화

        unitDataStore.forEach((unitData, unitName) => {
            // 1. Cycle별로 데이터 그룹화
            const cycles = {};
            unitData.tableData.forEach((rowData, index) => {
                // Cycle이 '0'이거나 비어있는 데이터는 분석에서 제외
                if (!rowData.cycle || rowData.cycle === '0') {
                    return;
                }

                const cycle = rowData.cycle;
                if (!cycles[cycle]) {
                    cycles[cycle] = [];
                }
                let endTime;
                if (index + 1 < unitData.tableData.length) {
                    endTime = unitData.tableData[index + 1].startTime;
                } else {
                    // 마지막 행은 동영상 전체 길이를 기준으로 함
                    // videoPlayer.duration이 유효한지 확인
                    const duration = unitData.videoSrc ? videoPlayer.duration : 0;
                    endTime = duration;
                }
                const interval = endTime - rowData.startTime;
                cycles[cycle].push({ valueType: rowData.valueType, interval: interval });
            });

            // 2. 각 Cycle의 VA, BVA, NVA 합계 계산 후, 전체 Cycle의 평균 계산
            const cycleAnalyses = { VA: [], BVA: [], NVA: [] };
            Object.values(cycles).forEach(cycleData => {
                const cycleTotals = { VA: 0, BVA: 0, NVA: 0 };
                cycleData.forEach(item => {
                    cycleTotals[item.valueType] += item.interval;
                });
                cycleAnalyses.VA.push(cycleTotals.VA);
                cycleAnalyses.BVA.push(cycleTotals.BVA);
                cycleAnalyses.NVA.push(cycleTotals.NVA);
            });

            const avgVA = cycleAnalyses.VA.length > 0 ? cycleAnalyses.VA.reduce((a, b) => a + b, 0) / cycleAnalyses.VA.length : 0;
            const avgBVA = cycleAnalyses.BVA.length > 0 ? cycleAnalyses.BVA.reduce((a, b) => a + b, 0) / cycleAnalyses.BVA.length : 0;
            const avgNVA = cycleAnalyses.NVA.length > 0 ? cycleAnalyses.NVA.reduce((a, b) => a + b, 0) / cycleAnalyses.NVA.length : 0;
            const totalAvg = avgVA + avgBVA + avgNVA;

            tableHtml += `
                <tr>
                    <td>${unitName}</td>
                    <td>${avgVA.toFixed(2)}</td>
                    <td>${avgBVA.toFixed(2)}</td>
                    <td>${avgNVA.toFixed(2)}</td>
                    <td>${totalAvg.toFixed(2)}</td>
                </tr>
            `;

            // 파이 차트 생성
            if (totalAvg > 0) {
                const chartWrapper = document.createElement('div');
                chartWrapper.className = 'chart-wrapper';

                const chartTitle = document.createElement('h3');
                chartTitle.textContent = `${unitName} Analysis`;

                const canvas = document.createElement('canvas');
                chartWrapper.appendChild(chartTitle);
                chartWrapper.appendChild(canvas);
                analysisChartContainer.appendChild(chartWrapper);

                Chart.register(ChartDataLabels); // 데이터 레이블 플러그인 등록
                new Chart(canvas, {
                    type: 'pie',
                    data: {
                        labels: ['VA', 'BVA', 'NVA'],
                        datasets: [{
                            label: 'Time (s)',
                            data: [avgVA, avgBVA, avgNVA],
                            backgroundColor: [
                                'rgba(40, 167, 69, 0.7)',  // Green for VA
                                'rgba(255, 193, 7, 0.7)',   // Yellow for BVA
                                'rgba(220, 53, 69, 0.7)'    // Red for NVA
                            ],
                            borderColor: [
                                'rgba(40, 167, 69, 1)',
                                'rgba(255, 193, 7, 1)',
                                'rgba(220, 53, 69, 1)'
                            ],
                            borderWidth: 1
                        }]
                    },
                    options: {
                        plugins: {
                            datalabels: {
                                // 2. 차트 위에 백분율 표시
                                formatter: (value, ctx) => {
                                    const sum = ctx.chart.data.datasets[0].data.reduce((a, b) => a + b, 0);
                                    const percentage = (value * 100 / sum).toFixed(1) + '%';
                                    return percentage;
                                },
                                color: '#fff',
                                font: { weight: 'bold' }
                            }
                        }
                    }
                });
            }
        });

        tableHtml += '</tbody></table>';
        analysisTableContainer.innerHTML = tableHtml;
    }

    // 동영상 클릭/더블클릭 시 기본 동작(재생/멈춤, 전체화면) 방지
    videoPlayer.addEventListener('click', (event) => {
        // 컨트롤 바를 클릭한 경우는 기본 동작 허용
        if (event.offsetY < videoPlayer.offsetHeight - 40) {
            event.preventDefault();
        }
    });
    videoPlayer.addEventListener('dblclick', (event) => {
        event.preventDefault(); // 전체화면 기본 동작 방지

        // 동영상이 로드되지 않았으면 아무 작업도 하지 않음
        if (!videoPlayer.src || !videoPlayer.duration) {
            alert('먼저 동영상을 불러오세요.');
            return;
        }

        const selectedUnitOption = unitSelector.options[unitSelector.selectedIndex];
        if (!selectedUnitOption) {
            alert('Unit을 선택하세요.');
            return;
        }

        const startTime = videoPlayer.currentTime;

        // 1. 중복 Start Time 확인
        const currentUnit = unitSelector.value;
        const currentUnitData = unitDataStore.get(currentUnit);

        if (currentUnitData.tableData.some(row => row.startTime === startTime)) {
            alert('동일한 시작 시간의 데이터가 이미 존재합니다.');
            return;
        }

        // 새 행 데이터 객체 생성
        const newRowData = {
            startTime: startTime,
            // Time Interval은 updateAllRowTimes에서 계산되므로 초기값 불필요
            process: 'Process',
            cycle: '0',
            workElement: 'Element',
            unitMotion: 'Motion',
            valueType: 'VA'
        };

        // 데이터 저장 및 정렬
        currentUnitData.tableData.push(newRowData);
        currentUnitData.tableData.sort((a, b) => a.startTime - b.startTime);

        // 테이블 다시 그리기
        renderTable(currentUnit);
    });

    // 테이블을 데이터 기준으로 다시 그리는 함수
    function renderTable(unitName) {
        const unitData = unitDataStore.get(unitName);
        if (!unitData) return;

        workTableBody.innerHTML = ''; // 테이블 비우기

        unitData.tableData.forEach(rowData => {
            const newRow = workTableBody.insertRow();
            newRow.dataset.startTime = rowData.startTime;

            // 각 셀(cell)에 데이터 추가
            newRow.insertCell(0).textContent = unitName;

            // Process, Work Element, Unit Motion
            const createInputCell = (value, key) => {
                const cell = document.createElement('td');
                const input = document.createElement('input');
                input.type = 'text';
                input.value = value;
                input.addEventListener('change', (e) => {
                    rowData[key] = e.target.value;
                });
                cell.appendChild(input);
                return cell;
            };
            newRow.appendChild(createInputCell(rowData.process, 'process'));

            // Cycle
            const createCycleCell = () => {
                const cell = document.createElement('td');
                const input = document.createElement('input');
                input.type = 'number';
                input.min = 0;
                input.max = 999;
                input.value = rowData.cycle;
                input.addEventListener('change', (e) => {
                    rowData.cycle = e.target.value;
                });
                cell.appendChild(input);
                return cell;
            };
            newRow.appendChild(createCycleCell());

            newRow.appendChild(createInputCell(rowData.workElement, 'workElement'));
            newRow.appendChild(createInputCell(rowData.unitMotion, 'unitMotion'));

            // Start Time, End Time, Time Interval (빈 셀 먼저 생성)
            newRow.insertCell(5);
            newRow.insertCell(6);
            newRow.insertCell(7); // Time Interval 셀

            // Value Type
            const createValueTypeCell = () => {
                const cell = document.createElement('td');
                const select = document.createElement('select');
                ['VA', 'BVA', 'NVA'].forEach(val => select.add(new Option(val, val)));
                select.value = rowData.valueType;
                select.addEventListener('change', (e) => {
                    rowData.valueType = e.target.value;
                    cell.className = `value-type-${e.target.value.toLowerCase()}`;
                });
                cell.className = `value-type-${rowData.valueType.toLowerCase()}`;
                cell.appendChild(select);
                return cell;
            };
            newRow.appendChild(createValueTypeCell());
        });

        updateAllRowTimes();
    }

    // Unit 변경 시 UI 업데이트 함수
    function switchUnit(unitName) {
        const unitData = unitDataStore.get(unitName);
        if (!unitData) {
            console.error(`${unitName}에 대한 데이터가 없습니다.`);
            return;
        }

        // 동영상 정보 복원
        videoPlayer.src = unitData.videoSrc || '';
        videoPathDisplay.textContent = unitData.videoName;

        // 테이블 복원
        renderTable(unitName);
    }
    
    // --- 동영상 드래그 앤 드롭 기능 ---
    videoSection.addEventListener('dragover', (event) => {
        event.preventDefault();
        videoSection.style.borderColor = '#007bff'; // 드래그 오버 시각적 피드백
    });

    videoSection.addEventListener('dragleave', () => {
        videoSection.style.borderColor = '#e0e0e0'; // 원래 테두리로 복원
    });

    videoSection.addEventListener('drop', (event) => {
        event.preventDefault();
        videoSection.style.borderColor = '#e0e0e0'; // 원래 테두리로 복원
        const file = event.dataTransfer.files[0];
        loadVideoFile(file);
    });

    // --- 동영상 포커스 및 키보드 이벤트 관리 ---
    document.addEventListener('click', (event) => {
        // 클릭된 위치가 video-section 내부인지 확인
        if (event.target.closest('.video-section')) {
            isVideoFocused = true;
        } else {
            isVideoFocused = false;
        }
    });

    // --- 테이블 행 상호작용 ---
    workTableBody.addEventListener('click', (event) => {
        const clickedRow = event.target.closest('tr');
        if (!clickedRow) return;

        // 다른 입력 요소가 아닌 행 자체를 클릭했을 때만 동작
        const targetTagName = event.target.tagName.toLowerCase();
        if (['input', 'select'].includes(targetTagName)) {
            return;
        } else {
            // 1. 행 클릭 시 Start Time으로 동영상 이동
            const startTime = parseFloat(clickedRow.dataset.startTime);
            if (!isNaN(startTime)) {
                videoPlayer.currentTime = startTime;
            }

            // 기존에 선택된 행이 있다면 하이라이트 제거
            const currentlySelected = workTableBody.querySelector('.row-selected');
            if (currentlySelected) {
                currentlySelected.classList.remove('row-selected');
            }
        }
        // 새로 클릭된 행에 하이라이트 추가
        clickedRow.classList.add('row-selected');
    });

    // 현재 시간 표시 기능
    const updateCurrentTime = () => {
        const currentTime = videoPlayer.currentTime.toFixed(2);
        const duration = videoPlayer.duration ? videoPlayer.duration.toFixed(2) : '0.00';
        timeDisplay.textContent = `${currentTime} / ${duration}`;
    };

    videoPlayer.addEventListener('timeupdate', updateCurrentTime);
    videoPlayer.addEventListener('loadedmetadata', () => {
        // 동영상 메타데이터가 로드된 후, 시간 관련 UI를 모두 업데이트
        updateCurrentTime();
        updateAllRowTimes();
    });

    // 초를 HH:MM:SS.ss 형식으로 변환하는 함수
    function formatTime(totalSeconds) {
        const hours = Math.floor(totalSeconds / 3600).toString().padStart(2, '0');
        const minutes = Math.floor((totalSeconds % 3600) / 60).toString().padStart(2, '0');
        const seconds = (totalSeconds % 60).toFixed(2).padStart(5, '0');
        return `${hours}:${minutes}:${seconds}`;
    }

    // 테이블의 모든 행의 EndTime과 Interval을 재계산하는 함수
    function updateAllRowTimes() {
        if (!videoPlayer.duration) return;

        const allRows = Array.from(workTableBody.querySelectorAll('tr'));

        for (let i = 0; i < allRows.length; i++) {
            const currentRow = allRows[i];
            const startTime = parseFloat(currentRow.dataset.startTime);

            let endTime;
            // 다음 행이 있으면 다음 행의 시작 시간을, 없으면 동영상 전체 길이를 사용
            if (i + 1 < allRows.length) { // 다음 행이 있으면
                endTime = parseFloat(allRows[i + 1].dataset.startTime);
            } else { // 마지막 행이면
                endTime = videoPlayer.duration;
            }

            const timeInterval = endTime - startTime;

            currentRow.dataset.endTime = endTime; // End Time도 데이터 속성으로 저장
            currentRow.cells[5].textContent = formatTime(startTime);
            currentRow.cells[6].textContent = formatTime(endTime);

            // Time Interval 셀을 숫자 입력으로 만들기
            const intervalCell = currentRow.cells[7];
            if (!intervalCell.querySelector('input')) {
                const input = document.createElement('input');
                input.type = 'number';
                input.step = 0.01;
                input.min = 0;
                input.addEventListener('change', (e) => {
                    handleIntervalChange(i, parseFloat(e.target.value));
                });
                // 1. 마우스 휠로 Time Interval 조절 기능 추가
                input.addEventListener('wheel', (e) => {
                    e.preventDefault(); // 페이지 스크롤 방지

                    let currentValue = parseFloat(input.value);

                    // 휠 방향에 따라 값 조정
                    if (e.deltaY < 0) { // 휠 위로
                        currentValue += 0.01;
                    } else { // 휠 아래로
                        currentValue -= 0.01;
                    }

                    input.value = Math.max(0, currentValue).toFixed(2); // 0 미만 방지

                    input.dispatchEvent(new Event('change')); // change 이벤트를 발생시켜 업데이트 로직 실행
                });
                intervalCell.innerHTML = ''; // 기존 내용 삭제
                intervalCell.appendChild(input);
            }
            intervalCell.querySelector('input').value = timeInterval.toFixed(2);
        }
    }

    // Time Interval 변경 시 후속 행들의 시간을 연쇄적으로 업데이트하는 함수
    function handleIntervalChange(changedRowIndex, newInterval) {
        const currentUnitData = unitDataStore.get(unitSelector.value);
        const tableData = currentUnitData.tableData;

        if (newInterval < 0) {
            newInterval = 0;
        }

        // 1. 현재 행의 EndTime 업데이트
        const currentRow = tableData[changedRowIndex];
        const newEndTime = currentRow.startTime + newInterval;

        // 2. 다음 행이 존재하면, 다음 행의 StartTime과 Interval을 조정
        if (changedRowIndex + 1 < tableData.length) {
            const nextRow = tableData[changedRowIndex + 1];
            const nextRowOriginalEndTime = nextRow.startTime + parseFloat(workTableBody.rows[changedRowIndex + 1].cells[7].querySelector('input').value);
            
            nextRow.startTime = newEndTime; // 다음 행의 StartTime을 현재 행의 새 EndTime으로 설정
            // 다음 행의 Interval은 (기존 EndTime - 새 StartTime)으로 재계산
        }

        // 3. 변경된 데이터로 테이블을 다시 그려서 모든 변경사항을 화면에 반영
        renderTable(unitSelector.value);

        // 4. 조절된 위치(새로운 End Time)로 동영상 이동
        videoPlayer.currentTime = newEndTime;
    }

    // 키보드 단축키 기능
    document.addEventListener('keydown', (event) => {
        // 입력 필드에 포커스가 있을 때는 단축키를 비활성화
        if (event.target.tagName === 'INPUT' || event.target.tagName === 'TEXTAREA') {
            return;
        }

        // 2. 'R' 키로 선택 구간 재생
        if (event.key.toLowerCase() === 'r') {
            event.preventDefault();
            const selectedRow = workTableBody.querySelector('.row-selected');
            if (selectedRow && selectedRow.dataset.startTime && selectedRow.dataset.endTime) {
                const startTime = parseFloat(selectedRow.dataset.startTime);
                const endTime = parseFloat(selectedRow.dataset.endTime);

                videoPlayer.currentTime = startTime;
                videoPlayer.play();

                const checkTime = () => {
                    if (videoPlayer.currentTime >= endTime) {
                        videoPlayer.pause();
                        videoPlayer.currentTime = endTime; // 정확히 End Time에 멈춤
                        videoPlayer.removeEventListener('timeupdate', checkTime);
                    }
                };

                // 기존 리스너를 제거하고 새로 추가하여 중복 방지
                videoPlayer.removeEventListener('timeupdate', checkTime);
                videoPlayer.addEventListener('timeupdate', checkTime);
            }
            return; // 다른 단축키와 겹치지 않도록 여기서 종료
        }

        const key = event.key.toLowerCase();
        let timeChange = 0;

        switch (key) {
            case 'q': timeChange = -1.0; break;
            case 'w': timeChange = 1.0; break;
            case 'a': timeChange = -0.1; break;
            case 's': timeChange = 0.1; break;
            case 'z': timeChange = -0.01; break;
            case 'x': timeChange = 0.01; break;
        }

        // 1. 동영상이 포커스되었을 때만 시간 이동 단축키 동작
        if (timeChange !== 0 && isVideoFocused) {
            event.preventDefault(); // 브라우저 기본 동작(예: 페이지 스크롤) 방지
            videoPlayer.currentTime = Math.max(0, videoPlayer.currentTime + timeChange);
        }
    });

    // 행 삭제 기능 (Delete 키)
    document.addEventListener('keydown', (event) => {
        if (event.key === 'Delete') {
            const selectedRow = workTableBody.querySelector('.row-selected');
            if (selectedRow) {
                event.preventDefault(); // 기본 동작 방지
                const startTimeToDelete = parseFloat(selectedRow.dataset.startTime);
                const currentUnit = unitSelector.value;
                const currentUnitData = unitDataStore.get(currentUnit);
                const indexToDelete = currentUnitData.tableData.findIndex(row => row.startTime === startTimeToDelete);
                if (indexToDelete > -1) {
                    currentUnitData.tableData.splice(indexToDelete, 1);
                    renderTable(currentUnit);
                }
            }
        }
    });

    // --- 화면 크기 조절 기능 ---
    const resizer = document.getElementById('resizer');
    const leftPanel = document.querySelector('.left-panel');
    const container = document.querySelector('.container');

    let isResizing = false;

    resizer.addEventListener('mousedown', (e) => {
        isResizing = true;
        document.body.style.cursor = 'col-resize'; // 전체 페이지 커서 변경
        document.body.style.userSelect = 'none'; // 드래그 중 텍스트 선택 방지
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        // container의 왼쪽 끝을 기준으로 한 마우스 위치
        const newLeftWidth = e.clientX - container.getBoundingClientRect().left;
        // 최소/최대 너비 제한
        if (newLeftWidth > 300 && newLeftWidth < window.innerWidth - 300) {
            container.style.gridTemplateColumns = `${newLeftWidth}px 1fr`;
        }
    });

    document.addEventListener('mouseup', () => {
        isResizing = false;
        document.body.style.cursor = 'default';
        document.body.style.userSelect = 'auto';
    });

    // --- 파일 저장 및 불러오기 기능 ---

    // 1. Excel 저장 기능
    saveExcelBtn.addEventListener('click', () => {
        if (unitDataStore.size === 0) {
            alert("저장할 Unit 데이터가 없습니다.");
            return;
        }

        const wb = XLSX.utils.book_new();
        const allDataForSummary = [];
        const headers = Array.from(document.querySelectorAll('.work-table th')).map(th => th.textContent);

        // 열 너비 자동 계산을 위한 헬퍼 함수
        const getColumnWidths = (data) => {
            const widths = [];
            data.forEach(row => {
                (row || []).forEach((cell, i) => {
                    const cellLength = cell ? String(cell).length : 0;
                    widths[i] = Math.max(widths[i] || 0, cellLength);
                });
            });
            return widths.map(w => ({ wch: w + 1 })); // 약간의 여백 추가
        };

        // 각 Unit을 별도의 시트로 저장
        unitDataStore.forEach((unitData, unitName) => {
            const sheetData = [];
            // 1. 동영상 경로 추가
            sheetData.push(['동영상 파일', unitData.videoName]);
            sheetData.push([]); // 빈 줄
            // 2. 헤더 추가
            sheetData.push(headers);

            // 3. 데이터 행 추가
            unitData.tableData.forEach((rowData, index) => {
                const { startTime } = rowData;
                let endTime = videoPlayer.duration; // 기본값

                // 다음 행의 startTime을 찾아서 endTime으로 설정
                if (index + 1 < unitData.tableData.length) {
                    endTime = unitData.tableData[index + 1].startTime;
                }

                const dataRow = [
                    unitName,
                    rowData.process,
                    rowData.cycle,
                    rowData.workElement,
                    rowData.unitMotion,
                    formatTime(startTime),
                    formatTime(endTime),
                    (endTime - startTime).toFixed(2),
                    rowData.valueType
                ];
                sheetData.push(dataRow);
                allDataForSummary.push(dataRow); // 요약 시트용 데이터 축적
            });

            const ws = XLSX.utils.aoa_to_sheet(sheetData);
            ws['!cols'] = getColumnWidths(sheetData); // 열 너비 적용
            XLSX.utils.book_append_sheet(wb, ws, unitName);
        });

        // 3. 마지막에 '분석 요약' 시트 추가
        if (allDataForSummary.length > 0) {
            const summarySheetData = [headers, ...allDataForSummary];
            const summary_ws = XLSX.utils.aoa_to_sheet(summarySheetData);
            summary_ws['!cols'] = getColumnWidths(summarySheetData); // 열 너비 적용
            XLSX.utils.book_append_sheet(wb, summary_ws, "분석 요약");
        }

        // 파일명 생성: TMO_첫번째유닛이름_날짜시간.xlsx
        const firstUnitName = unitDataStore.keys().next().value || 'data';
        const now = new Date();
        const timestamp = now.getFullYear().toString() +
                        (now.getMonth() + 1).toString().padStart(2, '0') +
                        now.getDate().toString().padStart(2, '0') +
                        now.getHours().toString().padStart(2, '0') +
                        now.getMinutes().toString().padStart(2, '0') +
                        now.getSeconds().toString().padStart(2, '0');
        const fileName = `TMO_${firstUnitName}_${timestamp}.xlsx`;

        XLSX.writeFile(wb, fileName);
    });

    // 2. Excel 불러오기 기능
    loadExcelBtn.addEventListener('click', () => {
        excelUpload.click();
    });

    excelUpload.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // 기존 데이터 초기화
                unitDataStore.clear();
                unitSelector.innerHTML = '';

                workbook.SheetNames.forEach(sheetName => {
                    if (sheetName === "분석 요약") return; // 요약 시트는 건너뜀

                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                    // Unit 생성 및 데이터 저장소 초기화
                    initializeUnitData(sheetName);
                    const newOption = new Option(sheetName, sheetName);
                    unitSelector.add(newOption);

                    const unitData = unitDataStore.get(sheetName);

                    // 동영상 경로 파싱
                    unitData.videoName = jsonData[0][1] || '선택된 파일 없음';

                    // 테이블 데이터 파싱 (헤더 제외)
                    const tableRows = jsonData.slice(3);
                    unitData.tableData = tableRows.filter(row => row && row[5] != null).map(row => { // 1. 유효한 행만 필터링
                        const timeString = String(row[5]);
                        if (timeString && timeString.includes(':')) {
                            const timeParts = timeString.split(':');
                            const startTimeSeconds = parseInt(timeParts[0]) * 3600 + parseInt(timeParts[1]) * 60 + parseFloat(timeParts[2]);
                            return {
                                startTime: startTimeSeconds,
                                process: row[1],
                                cycle: row[2],
                                workElement: row[3],
                                unitMotion: row[4],
                                valueType: row[8]
                            };
                        }
                        return null; // 유효하지 않은 행은 null 반환
                    }).sort((a, b) => a.startTime - b.startTime);
                });

                // 첫 번째 Unit으로 UI 전환
                if (unitSelector.options.length > 0) {
                    unitSelector.selectedIndex = 0;
                    switchUnit(unitSelector.value);
                }

            } catch (error) {
                console.error("Excel 파일 처리 중 오류 발생:", error);
                alert("Excel 파일을 불러오는 데 실패했습니다. 파일 형식을 확인해주세요.");
            }
        };
        reader.readAsArrayBuffer(file);
    });

    // Reset 버튼 기능
    const resetBtn = document.getElementById('resetBtn');
    resetBtn.addEventListener('click', () => {
        if (confirm('작업 전 초기상태로 변경하시겠습니까?')) {
            // 데이터 저장소 초기화
            unitDataStore.clear();
            // 콤보박스 초기화
            unitSelector.innerHTML = '';
            // 기본 Unit('Unit1') 다시 생성 및 UI 전환
            const defaultOption = new Option('Unit1', 'Unit1');
            unitSelector.add(defaultOption);
            initializeUnitData('Unit1');
            switchUnit('Unit1'); // 2. UI를 초기 상태로 즉시 업데이트
        }
    });
    // --- Unit 관리 기능 ---

    // 페이지 로드 시 기본 Unit('Unit1') 데이터 초기화
    const defaultOption = new Option('Unit1', 'Unit1');
    unitSelector.add(defaultOption);
    initializeUnitData('Unit1');
    
    // 초기 언어 설정
    updateLanguage(languageSelector.value);

    // Unit 선택 변경 시 이벤트
    unitSelector.addEventListener('change', (e) => {
        switchUnit(e.target.value);
    });

    // 팝업(모달) 열기 함수
    const openModal = (action, title, placeholder = '') => {
        currentModalAction = action;
        modalTitle.textContent = title;
        unitNameInput.value = placeholder;
        unitModal.style.display = 'flex';
        unitNameInput.focus();
    };

    // 팝업(모달) 닫기 함수
    const closeModal = () => {
        unitModal.style.display = 'none';
        unitNameInput.value = '';
    };

    // '추가' 버튼 클릭
    addUnitBtn.addEventListener('click', () => {
        openModal('add', '새 Unit 추가');
    });

    // '수정' 버튼 클릭
    editUnitBtn.addEventListener('click', () => {
        const selectedOption = unitSelector.options[unitSelector.selectedIndex];
        if (!selectedOption) {
            alert('수정할 Unit을 선택하세요.');
            return;
        }
        openModal('edit', 'Unit 이름 수정', selectedOption.textContent);
    });



    // '삭제' 버튼 클릭
    deleteUnitBtn.addEventListener('click', () => {
        const selectedOption = unitSelector.options[unitSelector.selectedIndex];
        if (!selectedOption) {
            alert('삭제할 Unit을 선택하세요.');
            return;
        }
        if (confirm(`'${selectedOption.textContent}' Unit과 모든 관련 데이터를 정말 삭제하시겠습니까?`)) {
            const unitNameToDelete = selectedOption.value;
            
            // 데이터 저장소에서 삭제
            unitDataStore.delete(unitNameToDelete);
            
            // 콤보박스에서 삭제
            unitSelector.remove(unitSelector.selectedIndex);

            // 만약 삭제 후 아무것도 없으면 기본값 다시 추가
            if (unitSelector.options.length === 0) {
                addUnitBtn.click(); // 새 Unit 추가 모달을 띄움
            }
            else { switchUnit(unitSelector.value); } // 남은 Unit 중 첫번째 것으로 화면 전환
        }
    });

    // 팝업의 '확인' 버튼 클릭
    modalOkBtn.addEventListener('click', () => {
        const unitName = unitNameInput.value.trim();
        if (!unitName) {
            alert('Unit 이름을 입력하세요.');
            return;
        }

        if (currentModalAction === 'add') {
            const newOption = new Option(unitName, unitName);
            initializeUnitData(unitName); // 새 Unit 데이터 공간 초기화
            unitSelector.add(newOption);
            unitSelector.value = unitName; // 새로 추가된 항목을 선택
            switchUnit(unitName); // 새 Unit 화면으로 전환
        } else if (currentModalAction === 'edit') {
            const selectedOption = unitSelector.options[unitSelector.selectedIndex];
            if (selectedOption) {
                const oldUnitName = selectedOption.value;
                const unitData = unitDataStore.get(oldUnitName);
                unitDataStore.delete(oldUnitName);
                unitDataStore.set(unitName, unitData);

                selectedOption.textContent = unitName;
                selectedOption.value = unitName;
                // 1. 이름 변경 시 테이블을 다시 그려서 Unit 이름을 즉시 업데이트
                renderTable(unitName);
            }
        }
        closeModal();
    });
    
    // 팝업의 '취소' 버튼 클릭
    modalCancelBtn.addEventListener('click', closeModal);

    // 팝업 외부 클릭 시 닫기
    window.addEventListener('click', (event) => {
        if (event.target === unitModal) {
            closeModal();
        }
        // 설명서 모달 닫기
        if (event.target === manualModal) {
            manualModal.style.display = 'none';
        }
    });

    // Unit 이름 입력 팝업에서 Enter 키로 확인 기능
    unitNameInput.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') {
            modalOkBtn.click(); // '확인' 버튼 클릭 이벤트 트리거
        }
    });
});
