// ========================================================
// MICROSOFT 365 POWER AUTOMATE FLOW ENDPOINTS
// ========================================================
const MS_POST_URL = "https://default7725528852984ea581aaa13e604c30.ac.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/17/workflows/6378a04da38d4ceca40b396ec77642c4/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PYXEHEyx3q-D81KB7tbaazfXYrgwlbdEx3nbM2Pk5eU";
const MS_GET_URL  = "https://default7725528852984ea581aaa13e604c30.ac.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/17/workflows/b66bbdb67be54eeb819390b27497a5dc/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=YyOV5aIndELK6T6d4Q3kM_O7zDCUQ2yFngrgMPrQlec";

let CHECKPOINTS = {};
let OFFICERS = [];

let activeCheckpoint = null;
let telemetryGPS = { lat: null, lon: null, accuracy: null, timestamp: null };
let capturedDefectPhotos = {};
let activePhotoQuestionId = null;
let cameraStream = null;
let qrScanInterval = null;
let isSyncing = false;
let isNetworkOnline = navigator.onLine;

window.addEventListener("DOMContentLoaded", () => {
    loadFromLocalStorage();
    fetchLiveConfiguration();
    checkActiveConnection();

    window.addEventListener("online", () => {
        checkActiveConnection();
        fetchLiveConfiguration();
        silentBackgroundSync();
    });

    window.addEventListener("offline", () => {
        setNetworkStatus(false);
    });

    setInterval(() => {
        if (navigator.onLine) checkActiveConnection();
    }, 15000);
});

function fetchLiveConfiguration() {
    if (!navigator.onLine || !MS_GET_URL.startsWith("http")) return;

    fetch(MS_GET_URL)
        .then(res => {
            if (!res.ok) throw new Error("Config fetch error");
            return res.json();
        })
        .then(data => {
            processRemoteConfig(data);
        })
        .catch(err => {
            console.warn("Using offline configuration cache:", err);
        });
}

function processRemoteConfig(data) {
    if (data.checkpoints && Array.isArray(data.checkpoints)) {
        CHECKPOINTS = {};
        data.checkpoints.forEach(row => {
            const code = String(row.Code || "").trim();
            if (code) {
                CHECKPOINTS[code] = {
                    code: code,
                    name: String(row.Name || "").trim(),
                    lat: parseFloat(row.Latitude),
                    lon: parseFloat(row.Longitude),
                    tolerance: parseFloat(row.Tolerance) || 15,
                    formCode: String(row.FormCode || "fen001").trim().toLowerCase()
                };
            }
        });
        localStorage.setItem("dp_checkpoints", JSON.stringify(CHECKPOINTS));
    }

    if (data.officers && Array.isArray(data.officers)) {
        OFFICERS = [];
        data.officers.forEach(row => {
            const name = String(row.Name || "").trim();
            const code = String(row.Code || "").trim();
            if (name || code) {
                OFFICERS.push({ name, code });
            }
        });
        localStorage.setItem("dp_officers", JSON.stringify(OFFICERS));
        populateOfficersDatalist(OFFICERS);
    }
}

function loadFromLocalStorage() {
    const cp = localStorage.getItem("dp_checkpoints");
    const off = localStorage.getItem("dp_officers");
    if (cp) CHECKPOINTS = JSON.parse(cp);
    if (off) {
        OFFICERS = JSON.parse(off);
        populateOfficersDatalist(OFFICERS);
    }
}

function checkActiveConnection() {
    if (!navigator.onLine) {
        setNetworkStatus(false);
        return;
    }

    fetch(`https://clients3.google.com/generate_204?_=${Date.now()}`, { mode: 'no-cors', cache: 'no-store' })
        .then(() => {
            setNetworkStatus(true);
            silentBackgroundSync();
        })
        .catch(() => {
            setNetworkStatus(false);
        });
}

function setNetworkStatus(online) {
    isNetworkOnline = online;
    const el = document.getElementById("net-status-text");
    if (online) {
        el.className = "net-online";
        el.innerText = "متصل";
    } else {
        el.className = "net-offline";
        el.innerText = "غير متصل";
    }
}

function getQueue() {
    const q = localStorage.getItem("dp_pending_patrols");
    return q ? JSON.parse(q) : [];
}

function saveToQueue(payload) {
    const queue = getQueue();
    queue.push(payload);
    localStorage.setItem("dp_pending_patrols", JSON.stringify(queue));
}

function populateOfficersDatalist(list) {
    const dl = document.getElementById("officers-datalist");
    dl.innerHTML = "";
    list.forEach(o => {
        const opt = document.createElement("option");
        opt.value = `${o.name} (${o.code})`;
        dl.appendChild(opt);
    });
}

function parseOfficerInput(raw) {
    let name = raw.trim();
    let code = "";
    const m = raw.match(/^(.*?)(?:\s*\((.*?)\))?$/);
    if (m) {
        name = m[1].trim();
        code = m[2] ? m[2].trim() : "";
    }
    return { name, code };
}

function calculateHaversine(lat1, lon1, lat2, lon2) {
    if (!lat1 || !lon1 || !lat2 || !lon2) return 999999;
    const R = 6371e3;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat/2) * Math.sin(dLat/2) +
        Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
        Math.sin(dLon/2) * Math.sin(dLon/2);
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function executeSingleButtonPatrol() {
    const officerRaw = document.getElementById("officer-input").value.trim();
    const feedback = document.getElementById("status-feedback");
    feedback.innerText = "";

    if (!officerRaw) {
        feedback.className = "msg danger-text";
        feedback.innerText = "⚠️ برجاء إدخال اسم أو كود مسؤول الدورية أولاً.";
        return;
    }

    const modal = document.getElementById("qr-modal");
    const video = document.getElementById("webcam");

    navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" }, audio: false })
        .then(stream => {
            cameraStream = stream;
            video.srcObject = stream;
            video.setAttribute("playsinline", "true");
            video.play();
            modal.style.display = "flex";
            qrScanInterval = setInterval(scanQRFrame, 200);
        })
        .catch(() => {
            feedback.className = "msg danger-text";
            feedback.innerText = "تعذر تشغيل الكاميرا. يرجى تفعيل إذن الكاميرا من إعدادات المتصفح.";
        });
}

function terminateScannerModal() {
    clearInterval(qrScanInterval);
    if (cameraStream) {
        cameraStream.getTracks().forEach(t => t.stop());
        cameraStream = null;
    }
    document.getElementById("qr-modal").style.display = "none";
}

function scanQRFrame() {
    const video = document.getElementById("webcam");
    const canvas = document.getElementById("qr-canvas");
    const ctx = canvas.getContext("2d", { willReadFrequently: true });

    if (video.readyState === video.HAVE_ENOUGH_DATA) {
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

        const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        /* global jsQR */
        const code = jsQR(imgData.data, imgData.width, imgData.height, { inversionAttempts: "dontInvert" });

        if (code) {
            const scannedStr = code.data.trim();
            terminateScannerModal();

            activeCheckpoint = CHECKPOINTS[scannedStr] || {
                code: scannedStr,
                name: `نقطة غير مسجلة (${scannedStr})`,
                lat: 0,
                lon: 0,
                tolerance: 15,
                formCode: "fen001"
            };

            acquireLocationSilentlyAndShowForm();
        }
    }
}

function acquireLocationSilentlyAndShowForm() {
    const feedback = document.getElementById("status-feedback");
    feedback.className = "msg";
    feedback.innerText = "جاري فتح استمارة التفتيش...";

    if (!navigator.geolocation) {
        telemetryGPS = { lat: 0, lon: 0, accuracy: 9999, timestamp: Date.now() };
        displayInspectionForm();
        return;
    }

    navigator.geolocation.getCurrentPosition(
        (pos) => {
            telemetryGPS = {
                lat: pos.coords.latitude,
                lon: pos.coords.longitude,
                accuracy: Math.round(pos.coords.accuracy),
                timestamp: pos.timestamp || Date.now()
            };
            feedback.innerText = "";
            displayInspectionForm();
        },
        () => {
            telemetryGPS = { lat: 0, lon: 0, accuracy: 9999, timestamp: Date.now() };
            feedback.innerText = "";
            displayInspectionForm();
        },
        { enableHighAccuracy: true, timeout: 5000, maximumAge: 0 }
    );
}

function displayInspectionForm() {
    const formCode = (activeCheckpoint.formCode || "fen001").toString().toLowerCase().trim();
    const container = document.getElementById("dynamic-form-fields");
    const titleEl = document.getElementById("form-modal-title");
    container.innerHTML = "";
    capturedDefectPhotos = {};

    // FORM A: fen001 (Fence Inspection)
    if (formCode === "fen001") {
        titleEl.innerText = `استمارة فحص السور - ${activeCheckpoint.name}`;
        container.innerHTML = `
            <div class="form-group">
                <label>حالة السور:</label>
                <select name="fence_status" onchange="evaluatePhotoCondition('fence_status', this.value === 'Broken')">
                    <option value="Intact">سليم</option>
                    <option value="Broken">مكسور</option>
                </select>
                <div id="box_fence_status" class="conditional-box">
                    <label style="color:var(--dp-warning);">⚠️ مطلوب صورة توثيقية للكسر:</label>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('fence_status')">📷 التقاط صورة العطل</button>
                    <img id="preview_fence_status" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <div class="form-group">
                <label>حالة الإضاءة:</label>
                <select name="light_status" onchange="evaluatePhotoCondition('light_status', this.value === 'Defective')">
                    <option value="Working">سليمة</option>
                    <option value="Defective">غير سليمة</option>
                </select>
                <div id="box_light_status" class="conditional-box">
                    <label style="color:var(--dp-warning);">⚠️ مطلوب صورة توثيقية لخلل الإضاءة:</label>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('light_status')">📷 التقاط صورة الخلل</button>
                    <img id="preview_light_status" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <div class="form-group">
                <label>هل هناك نقطة اختراق في السور؟</label>
                <select name="breach_status" onchange="evaluatePhotoCondition('breach_status', this.value === 'Yes')">
                    <option value="No">لا</option>
                    <option value="Yes">نعم</option>
                </select>
                <div id="box_breach_status" class="conditional-box">
                    <label style="color:var(--dp-warning);">⚠️ مطلوب صورة توثيقية لنقطة الاختراق:</label>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('breach_status')">📷 التقاط صورة الاختراق</button>
                    <img id="preview_breach_status" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <div class="form-group">
                <label>ملاحظات إضافية:</label>
                <textarea name="notes" rows="2" placeholder="اكتب الملاحظات إن وجدت..."></textarea>
            </div>
        `;
    }
    // FORM B: adof001 (Office Inspection)
    else if (formCode === "adof001") {
        titleEl.innerText = `استمارة فحص المكاتب - ${activeCheckpoint.name}`;
        container.innerHTML = `
            <div class="form-group">
                <label>حالة المكتب:</label>
                <select name="office_status" onchange="evaluatePhotoCondition('office_status', this.value === 'Open')">
                    <option value="Closed">مغلق</option>
                    <option value="Open">مفتوح</option>
                </select>
                <div id="box_office_status" class="conditional-box">
                    <label>الإجراء التصحيحي لحالة المكتب:</label>
                    <select name="office_action">
                        <option value="Leave as is">البقاء كما هو عليه</option>
                        <option value="Monitor via cameras">متابعة من خلال الكاميرات</option>
                        <option value="Notify office owner">إبلاغ صاحب المكتب</option>
                        <option value="Closed by officer">تم الإغلاق</option>
                    </select>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('office_status')">📷 التقاط صورة توثيقية للمكتب</button>
                    <img id="preview_office_status" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <div class="form-group">
                <label>حالة التكييف:</label>
                <select name="ac_status" onchange="evaluatePhotoCondition('ac_status', this.value === 'Running')">
                    <option value="Closed">مغلق</option>
                    <option value="Running">مفتوح</option>
                </select>
                <div id="box_ac_status" class="conditional-box">
                    <label>الإجراء التصحيحي للتكييف:</label>
                    <select name="ac_action">
                        <option value="Leave as is">البقاء كما هو عليه</option>
                        <option value="Notify office owner">إبلاغ صاحب المكتب</option>
                        <option value="Closed by officer">تم الإغلاق</option>
                    </select>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('ac_status')">📷 التقاط صورة للتكييف</button>
                    <img id="preview_ac_status" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <div class="form-group">
                <label>هل هناك مقتنيات معرضة للسرقة؟</label>
                <select name="belongings_status" onchange="evaluatePhotoCondition('belongings_status', this.value === 'Yes')">
                    <option value="No">لا</option>
                    <option value="Yes">نعم</option>
                </select>
                <div id="box_belongings_status" class="conditional-box">
                    <label>الإجراء التصحيحي للمقتنيات:</label>
                    <select name="belongings_action">
                        <option value="Leave as is">البقاء كما هو عليه</option>
                        <option value="Monitor via cameras">متابعة من خلال الكاميرات</option>
                        <option value="Notify office owner">إبلاغ صاحب المكتب</option>
                        <option value="Handed over to control room">تم التسليم إلي غرفة التحكم</option>
                    </select>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('belongings_status')">📷 التقاط صورة المقتنيات</button>
                    <img id="preview_belongings_status" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <div class="form-group">
                <label>ملاحظات إضافية:</label>
                <textarea name="notes" rows="2" placeholder="اكتب الملاحظات إن وجدت..."></textarea>
            </div>
        `;
    }
    // FORM C: sp001 (Security Posts Inspection)
    else if (formCode === "sp001") {
        titleEl.innerText = `استمارة نقطة الحراسة - ${activeCheckpoint.name}`;
        container.innerHTML = `
            <div class="form-group">
                <label>فرد الأمن المتواجد بالنقطة:</label>
                <input type="text" name="guard_at_post" list="officers-datalist" placeholder="ابحث بالاسم أو الكود (مثال: 30061)..." autocomplete="off" required>
            </div>

            <!-- 1. Perimeter Cleanliness -->
            <div class="form-group">
                <label>حالة النطاق الأمني:</label>
                <select name="perimeter_clean" onchange="evaluatePhotoCondition('perimeter_photo', this.value === 'Dirty')">
                    <option value="Clean">نظيف</option>
                    <option value="Dirty">غير نظيف</option>
                </select>
                <div id="box_perimeter_photo" class="conditional-box">
                    <label style="color:var(--dp-warning);">⚠️ مطلوب صورة توثيقية للنطاق غير النظيف:</label>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('perimeter_photo')">📷 التقاط صورة النطاق</button>
                    <img id="preview_perimeter_photo" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <!-- 2. Safety Equipment -->
            <div class="form-group">
                <label>حالة أدوات السلامة:</label>
                <select name="safety_equip" onchange="evaluatePhotoCondition('equip_photo', this.value === 'Damaged')">
                    <option value="Intact">سليمة</option>
                    <option value="Damaged">غير سليمة</option>
                </select>
                <div id="box_equip_photo" class="conditional-box">
                    <label style="color:var(--dp-warning);">⚠️ مطلوب صورة توثيقية لتلف أدوات السلامة:</label>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('equip_photo')">📷 التقاط صورة الأدوات</button>
                    <img id="preview_equip_photo" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <!-- 3. Personnel Presence (Mandatory Photo on BOTH Present & Absent) -->
            <div class="form-group">
                <label>حالة فرد الأمن:</label>
                <select name="personnel_present" id="sp_personnel_present" onchange="handlePersonnelPresenceChange(this.value)">
                    <option value="Present">متواجد</option>
                    <option value="Absent">غير متواجد</option>
                </select>
                <div id="box_personnel_photo" class="conditional-box" style="display:block; background:#f0fdf4; border-color:var(--dp-green);">
                    <label id="lbl_personnel_photo" style="color:var(--dp-navy);">📷 مطلوب صورة توثيقية لإثبات تواجد فرد الأمن بالنقطة:</label>
                    <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('personnel_photo')">📷 التقاط صورة التواجد</button>
                    <img id="preview_personnel_photo" class="photo-preview" alt="معاينة الصورة">
                </div>
            </div>

            <!-- 4. Dependent Guard Behavior Fields -->
            <div id="guard_behavior_fields">
                <div class="form-group">
                    <label>هل فرد الأمن نائم؟</label>
                    <select name="is_sleeping" onchange="evaluatePhotoCondition('sleeping_photo', this.value === 'Yes')">
                        <option value="No">لا</option>
                        <option value="Yes">نعم</option>
                    </select>
                    <div id="box_sleeping_photo" class="conditional-box">
                        <label style="color:var(--danger);">⚠️ توثيق مخالفة نوم أثناء الخدمة:</label>
                        <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('sleeping_photo')">📷 التقاط صورة المخالفة</button>
                        <img id="preview_sleeping_photo" class="photo-preview" alt="معاينة الصورة">
                    </div>
                </div>

                <div class="form-group">
                    <label>هل فرد الأمن يرتدي الزي؟</label>
                    <select name="wearing_uniform" id="sp_wearing_uniform" onchange="handleWearingUniformChange(this.value)">
                        <option value="Yes">نعم</option>
                        <option value="No">لا</option>
                    </select>
                    <div id="box_uniform_photo" class="conditional-box">
                        <label style="color:var(--danger);">⚠️ توثيق مخالفة عدم ارتداء الزي:</label>
                        <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('uniform_photo')">📷 التقاط صورة المخالفة</button>
                        <img id="preview_uniform_photo" class="photo-preview" alt="معاينة الصورة">
                    </div>
                </div>

                <!-- 5. Dependent Neatness Field -->
                <div class="form-group" id="box_uniform_neat_container">
                    <label>هل الزي مهندم؟</label>
                    <select name="uniform_neat" onchange="evaluatePhotoCondition('neat_photo', this.value === 'No')">
                        <option value="Yes">نعم</option>
                        <option value="No">لا</option>
                    </select>
                    <div id="box_neat_photo" class="conditional-box">
                        <label style="color:var(--danger);">⚠️ توثيق مظهر الزي غير المهندم:</label>
                        <button type="button" class="btn btn-capture-sub" onclick="openPhotoCaptureModal('neat_photo')">📷 التقاط صورة الزي</button>
                        <img id="preview_neat_photo" class="photo-preview" alt="معاينة الصورة">
                    </div>
                </div>
            </div>

            <div class="form-group">
                <label>ملاحظات:</label>
                <textarea name="notes" rows="2" placeholder="اكتب الملاحظات إن وجدت..."></textarea>
            </div>
        `;
    }

    document.getElementById("form-modal").style.display = "flex";
}

function handlePersonnelPresenceChange(status) {
    const behaviorBox = document.getElementById("guard_behavior_fields");
    const photoBox = document.getElementById("box_personnel_photo");
    const photoLabel = document.getElementById("lbl_personnel_photo");

    if (status === "Absent") {
        if (behaviorBox) behaviorBox.style.display = "none";

        delete capturedDefectPhotos['sleeping_photo'];
        delete capturedDefectPhotos['uniform_photo'];
        delete capturedDefectPhotos['neat_photo'];

        if (photoBox) {
            photoBox.style.background = "#fff1f2";
            photoBox.style.borderColor = "var(--danger)";
        }
        if (photoLabel) {
            photoLabel.style.color = "var(--danger)";
            photoLabel.innerText = "⚠️ مطلوب صورة توثيقية لإثبات غياب فرد الأمن / خلو النقطة:";
        }
    } else {
        if (behaviorBox) behaviorBox.style.display = "block";

        if (photoBox) {
            photoBox.style.background = "#f0fdf4";
            photoBox.style.borderColor = "var(--dp-green)";
        }
        if (photoLabel) {
            photoLabel.style.color = "var(--dp-navy)";
            photoLabel.innerText = "📷 مطلوب صورة توثيقية لإثبات تواجد فرد الأمن بالنقطة:";
        }

        const uniformSelect = document.getElementById("sp_wearing_uniform");
        if (uniformSelect) {
            handleWearingUniformChange(uniformSelect.value);
        }
    }
}

function handleWearingUniformChange(value) {
    evaluatePhotoCondition('uniform_photo', value === 'No');

    const neatContainer = document.getElementById("box_uniform_neat_container");
    if (value === "No") {
        if (neatContainer) neatContainer.style.display = "none";
        delete capturedDefectPhotos['neat_photo'];
    } else {
        if (neatContainer) neatContainer.style.display = "block";
    }
}

function evaluatePhotoCondition(fieldId, isNegative) {
    const box = document.getElementById(`box_${fieldId}`);
    if (box) box.style.display = isNegative ? "block" : "none";
    if (!isNegative) delete capturedDefectPhotos[fieldId];
}

function closeFormModal() {
    document.getElementById("form-modal").style.display = "none";
}

let photoStream = null;
function openPhotoCaptureModal(questionId) {
    activePhotoQuestionId = questionId;
    const modal = document.getElementById("photo-capture-modal");
    const video = document.getElementById("photo-webcam");

    navigator.mediaDevices.getUserMedia({ video: { facingMode: "environment" }, audio: false })
        .then(stream => {
            photoStream = stream;
            video.srcObject = stream;
            video.setAttribute("playsinline", "true");
            video.play();
            modal.style.display = "flex";
        })
        .catch(() => alert("تعذر فتح كاميرا التوثيق."));
}

function closePhotoModal() {
    if (photoStream) {
        photoStream.getTracks().forEach(t => t.stop());
        photoStream = null;
    }
    document.getElementById("photo-capture-modal").style.display = "none";
}

function takeDefectSnapshot() {
    const video = document.getElementById("photo-webcam");
    const canvas = document.getElementById("photo-canvas");
    const ctx = canvas.getContext("2d", { willReadFrequently: true });

    canvas.width = video.videoWidth || 640;
    canvas.height = video.videoHeight || 480;
    ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

    const base64Data = canvas.toDataURL("image/jpeg", 0.5);
    capturedDefectPhotos[activePhotoQuestionId] = base64Data;

    const preview = document.getElementById(`preview_${activePhotoQuestionId}`);
    if (preview) {
        preview.src = base64Data;
        preview.style.display = "block";
    }

    closePhotoModal();
}

// ----------------------------------------------------
// PREFLIGHT-FREE SUBMISSION & SILENT AUTO-SYNC
// ----------------------------------------------------
function processInspectionSubmit(event) {
    event.preventDefault();
    const submitBtn = document.getElementById("submit-inspection-btn");
    const officerRaw = document.getElementById("officer-input").value.trim();

    submitBtn.disabled = true;
    submitBtn.innerText = "جاري ترحيل التقرير...";

    const parsedOfficer = parseOfficerInput(officerRaw);
    const formEl = event.target;
    const formData = new FormData(formEl);
    let answersObj = {};

    for (let [key, val] of formData.entries()) {
        answersObj[key] = val;
    }

    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    const captureDate = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const captureTime = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

    const distanceMeters = Math.round(calculateHaversine(
        telemetryGPS.lat, telemetryGPS.lon, activeCheckpoint.lat, activeCheckpoint.lon
    ));

    // Case-insensitive, robust defect detection across all 3 forms
    const currentForm = (activeCheckpoint.formCode || "fen001").toString().toLowerCase().trim();
    let isDefect = false;

    if (currentForm === "fen001") {
        const fenceVal = String(answersObj["fence_status"] || "").trim().toLowerCase();
        const lightVal = String(answersObj["light_status"] || "").trim().toLowerCase();
        const breachVal = String(answersObj["breach_status"] || "").trim().toLowerCase();
        isDefect = (fenceVal === "broken" || lightVal === "defective" || breachVal === "yes");
    } else if (currentForm === "adof001") {
        const officeVal = String(answersObj["office_status"] || "").trim().toLowerCase();
        const acVal = String(answersObj["ac_status"] || "").trim().toLowerCase();
        const belongVal = String(answersObj["belongings_status"] || "").trim().toLowerCase();
        isDefect = (officeVal === "open" || acVal === "running" || belongVal === "yes");
    } else if (currentForm === "sp001") {
        const perimVal = String(answersObj["perimeter_clean"] || "").trim().toLowerCase();
        const personVal = String(answersObj["personnel_present"] || "").trim().toLowerCase();
        const equipVal = String(answersObj["safety_equip"] || "").trim().toLowerCase();

        if (personVal === "absent") {
            isDefect = true;
            answersObj["is_sleeping"] = "N/A";
            answersObj["wearing_uniform"] = "N/A";
            answersObj["uniform_neat"] = "N/A";
        } else {
            const sleepVal = String(answersObj["is_sleeping"] || "").trim().toLowerCase();
            const unifVal = String(answersObj["wearing_uniform"] || "").trim().toLowerCase();

            if (unifVal === "no") {
                answersObj["uniform_neat"] = "N/A";
            }
            const neatVal = String(answersObj["uniform_neat"] || "").trim().toLowerCase();

            isDefect = (perimVal === "dirty" || equipVal === "damaged" || sleepVal === "yes" || unifVal === "no" || neatVal === "no");
        }
    }

    const allowedTolerance = Math.max(activeCheckpoint.tolerance || 15, telemetryGPS.accuracy || 0);
    const isSuspicious = distanceMeters > allowedTolerance;

    // Construct Direct Google Maps Navigation URL
    const mapsLink = (telemetryGPS.lat && telemetryGPS.lat !== 0)
        ? `https://www.google.com/maps?q=${telemetryGPS.lat},${telemetryGPS.lon}`
        : "No Location";

    const payload = {
        officerName: parsedOfficer.name,
        officerCode: parsedOfficer.code,
        checkpointCode: activeCheckpoint.code,
        checkpointName: activeCheckpoint.name,
        formCode: activeCheckpoint.formCode || "fen001",
        answers: answersObj,
        photos: capturedDefectPhotos,
        captureDate: captureDate,
        captureTime: captureTime,
        internetStatus: isNetworkOnline ? "Online" : "Offline",
        latitude: telemetryGPS.lat,
        longitude: telemetryGPS.lon,
        distance: distanceMeters,
        mapsUrl: mapsLink, // Direct Google Maps Link
        inspectionStatus: isDefect ? "FLAGGED_DEFECT" : "NORMAL",
        securityStatus: isSuspicious ? "SUSPICIOUS / MANIPULATED" : "VERIFIED",
        securityFlags: isSuspicious ? `OUT_OF_BOUNDS (${distanceMeters}m > ${allowedTolerance}m)` : "CLEAN"
    };

    if (!isNetworkOnline) {
        saveToQueue(payload);
        finishSubmissionFlow(formEl, submitBtn, `✔ تم تسجيل الفحص بنجاح!`);
        return;
    }

    fetch(MS_POST_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
    })
        .then(response => {
            if (!response.ok) throw new Error("HTTP error " + response.status);
            finishSubmissionFlow(formEl, submitBtn, `✔ تم تسجيل الفحص بنجاح!`);
        })
        .catch(err => {
            console.warn("Direct submission failed, adding to offline queue:", err);
            saveToQueue(payload);
            finishSubmissionFlow(formEl, submitBtn, `✔ تم تسجيل الفحص بنجاح!`);
        });
}

function finishSubmissionFlow(formEl, submitBtn, msgText) {
    const feedback = document.getElementById("status-feedback");
    closeFormModal();
    formEl.reset();
    submitBtn.disabled = false;
    submitBtn.innerText = "✔ إرسال تقرير الفحص";

    feedback.className = "msg success-text";
    feedback.innerText = msgText;

    activeCheckpoint = null;
    telemetryGPS = { lat: null, lon: null, accuracy: null, timestamp: null };
    capturedDefectPhotos = {};

    setTimeout(() => {
        if (feedback.innerText.includes("تم")) feedback.innerText = "";
    }, 5000);
}

// SILENT AUTO-SYNC
function silentBackgroundSync() {
    if (isSyncing || !isNetworkOnline) return;
    const queue = getQueue();
    if (queue.length === 0) return;

    isSyncing = true;

    function sendNext() {
        const currentQueue = getQueue();
        if (currentQueue.length === 0) {
            isSyncing = false;
            return;
        }

        const item = currentQueue[0];

        fetch(MS_POST_URL, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(item)
        })
            .then(res => {
                if (!res.ok) throw new Error("Sync status: " + res.status);
                const updatedQueue = getQueue();
                updatedQueue.shift();
                localStorage.setItem("dp_pending_patrols", JSON.stringify(updatedQueue));
                sendNext();
            })
            .catch(err => {
                console.warn("Sync failed, queue preserved:", err);
                isSyncing = false;
            });
    }

    sendNext();
}
