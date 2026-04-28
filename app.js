// PBHRC Social Media Dashboard - Core Logic

let globalRawData = [];
let dashboardData = initDashboardObject();

// Chart Instances
let stateCharts = {
    pieBroad: null,
    piePlatform: null,
    barNewOld: null,
    timeline: null,
    socialConversionTimeline: null,
    genderPie: null,
    countryBar: null,
    stateBar: null,
    platformConversion: null,
    platformDistribution: null,
    genderPlatformPie: null
};

Chart.defaults.color = document.documentElement.getAttribute('data-theme') === 'light' ? '#4a6175' : '#7e95ae';
Chart.defaults.font.family = "'Inter', sans-serif";
Chart.defaults.borderColor = document.documentElement.getAttribute('data-theme') === 'light' ? 'rgba(0,0,0,0.06)' : getGridColor();

function initDashboardObject() {
    return {
        totalPatients: 0,
        socialPatients: 0,
        socialNewPatients: 0,
        broadSource: {},
        socialPlatform: {},
        socialNewVsOld: { new: 0, old: 0 },
        socialDays: { 'Sunday': 0, 'Monday': 0, 'Tuesday': 0, 'Wednesday': 0, 'Thursday': 0, 'Friday': 0, 'Saturday': 0 },
        socialStates: {},
        platformConversion: {},
        platformTimeline: {},
        platformGender: {},
        timeline: {},
        socialConversionTimeline: {},
        demographics: {
            gender: {},
            country: {},
            state: {}
        }
    };
}

// ============ INITIALIZATION ============
document.addEventListener('DOMContentLoaded', function() {
    // Tab switching
    var navItems = document.querySelectorAll('.nav-item');
    for (var i = 0; i < navItems.length; i++) {
        navItems[i].addEventListener('click', function(e) {
            var allNav = document.querySelectorAll('.nav-item');
            var allTabs = document.querySelectorAll('.tab-pane');
            for (var j = 0; j < allNav.length; j++) allNav[j].classList.remove('active');
            for (var j = 0; j < allTabs.length; j++) {
                allTabs[j].classList.remove('active');
                allTabs[j].style.display = 'none';
            }
            this.classList.add('active');
            var targetTab = this.getAttribute('data-tab');
            var tabEl = document.getElementById(targetTab);
            if (tabEl) {
                tabEl.classList.add('active');
                tabEl.style.display = '';
            }
        });
    }

    // Sidebar Toggle
    var sidebarToggle = document.getElementById('sidebar-toggle');
    if(sidebarToggle) {
        sidebarToggle.addEventListener('click', function() {
            var sidebar = document.querySelector('.sidebar');
            if(sidebar) sidebar.classList.toggle('collapsed');
        });
    }

    // Time filter
    document.getElementById('timeframe-filter').addEventListener('change', function() {
        if (globalRawData.length > 0) {
            recalculateData();
        }
    });

    setupFileHandlers();
    setupLiveSync();
});

function setupLiveSync() {
    var modal = document.getElementById('sync-modal');
    if(!modal) return;
    
    var btnSync = document.getElementById('btn-live-sync');
    var btnSyncLg = document.getElementById('btn-live-sync-lg');
    var btnClose = document.getElementById('close-sync-modal');
    var btnSave = document.getElementById('btn-save-sync');
    var btnClear = document.getElementById('btn-clear-url');
    var inputUrl = document.getElementById('google-sheet-url');

    function openModal() {
        modal.classList.remove('hidden');
        inputUrl.value = localStorage.getItem('google_sheet_url') || '';
    }

    function closeModal() {
        modal.classList.add('hidden');
    }

    if(btnSync) btnSync.addEventListener('click', openModal);
    if(btnSyncLg) btnSyncLg.addEventListener('click', openModal);
    if(btnClose) btnClose.addEventListener('click', closeModal);

    if(btnSave) {
        btnSave.addEventListener('click', function() {
            var url = inputUrl.value.trim();
            if (!url) {
                alert('Please enter a valid URL.');
                return;
            }
            localStorage.setItem('google_sheet_url', url);
            closeModal();
            fetchGoogleSheet(url);
        });
    }

    if(btnClear) {
        btnClear.addEventListener('click', function() {
            localStorage.removeItem('google_sheet_url');
            inputUrl.value = '';
            alert('Saved URL cleared.');
        });
    }

    // Auto-sync on startup if URL exists
    var savedUrl = localStorage.getItem('google_sheet_url');
    if (savedUrl) {
        fetchGoogleSheet(savedUrl);
    }
}

function setupFileHandlers() {
    var splash = document.querySelector('.splash-content');
    
    // Prevent browser defaults on drag events
    var events = ['dragenter', 'dragover', 'dragleave', 'drop'];
    for (var i = 0; i < events.length; i++) {
        splash.addEventListener(events[i], function(e) { e.preventDefault(); e.stopPropagation(); }, false);
        document.body.addEventListener(events[i], function(e) { e.preventDefault(); e.stopPropagation(); }, false);
    }

    splash.addEventListener('dragenter', function() { splash.classList.add('dragover'); }, false);
    splash.addEventListener('dragover', function() { splash.classList.add('dragover'); }, false);
    splash.addEventListener('dragleave', function() { splash.classList.remove('dragover'); }, false);
    splash.addEventListener('drop', function(e) {
        splash.classList.remove('dragover');
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) {
            processFile(e.dataTransfer.files[0]);
        }
    }, false);

    document.getElementById('excel-upload').addEventListener('change', function(e) {
        if (e.target.files.length) {
            processFile(e.target.files[0]);
            e.target.value = '';
        }
    });
    document.getElementById('excel-upload-lg').addEventListener('change', function(e) {
        if (e.target.files.length) {
            processFile(e.target.files[0]);
            e.target.value = '';
        }
    });
}

// ============ SOURCE CLASSIFICATION ============
function classifyBroadSource(s) {
    if (!s || typeof s !== 'string') return 'Unknown';
    var lower = s.toLowerCase().trim();
    if (!lower) return 'Unknown';

    // Social Media
    if (lower.indexOf('socialmedia') !== -1 || lower.indexOf('social media') !== -1 ||
        lower.indexOf('facebook') !== -1 || lower.indexOf('instagram') !== -1 ||
        lower.indexOf('youtube') !== -1 || lower.indexOf("what's up") !== -1 ||
        lower.indexOf('whatsapp') !== -1 || lower.indexOf('insta') !== -1) {
        return 'Social Media';
    }
    // Website & Search
    if (lower.indexOf('website') !== -1 || lower.indexOf('google') !== -1 || lower.indexOf('search') !== -1) {
        return 'Website & Search';
    }
    // Existing Patient
    if (lower.indexOf('old') !== -1 || lower.indexOf('existing') !== -1 || lower.indexOf('already') !== -1 ||
        lower.indexOf('regular') !== -1 || lower.indexOf('previous') !== -1 || lower.indexOf('since') !== -1 ||
        lower.indexOf('follow up') !== -1 || lower.indexOf('visited') !== -1) {
        return 'Existing Patient';
    }
    // Referrals
    if (lower.indexOf('family') !== -1 || lower.indexOf('friend') !== -1 || lower.indexOf('relative') !== -1 ||
        lower.indexOf('refer') !== -1 || lower.indexOf('brother') !== -1 || lower.indexOf('mother') !== -1 ||
        lower.indexOf('father') !== -1 || lower.indexOf('husband') !== -1 || lower.indexOf('wife') !== -1 ||
        lower.indexOf('colleague') !== -1) {
        return 'Referrals';
    }
    // Doctor / Protocol
    if (lower.indexOf('doctor') !== -1 || lower.indexOf('dr.') !== -1 || lower.indexOf('protocol') !== -1) {
        return 'Doctor / Protocol';
    }
    // Ads
    if (lower.indexOf('banner') !== -1 || lower.indexOf('newspaper') !== -1 || lower.indexOf('ad') !== -1) {
        return 'Ads & Print';
    }
    return 'Other / Direct';
}

function classifySocialPlatform(s) {
    if (!s || typeof s !== 'string') return null;
    var lower = s.toLowerCase().trim();

    var fb = lower.indexOf('facebook') !== -1;
    var ig = lower.indexOf('instagram') !== -1 || lower.indexOf('insta') !== -1;
    var yt = lower.indexOf('youtube') !== -1;
    var wa = lower.indexOf('whatsapp') !== -1 || lower.indexOf("what's up") !== -1;
    var li = lower.indexOf('linkedin') !== -1;

    // Prefer the most specific match
    if (li) return 'LinkedIn';
    if (yt) return 'YouTube';
    if (wa) return 'WhatsApp';
    if (fb && ig) return 'Facebook & Instagram';
    if (fb) return 'Facebook';
    if (ig) return 'Instagram';

    // Generic "social media" with no platform named
    return 'Generic Social Media';
}

// ============ COUNTRY / STATE NORMALISATION ============
function normalizeCountry(raw) {
    if (!raw) return null;
    var s = raw.trim().toLowerCase();
    if (!s) return null;

    // India variants (English + Transliterated + Bengali ভারত + Hindi भारत)
    if (s === 'india' || s === 'indian' || s === 'bharat' || s === 'bharath' ||
        s === '\u09ad\u09be\u09b0\u09a4' || // ভারত (Bengali)
        s === '\u092d\u093e\u09b0\u09a4' || // भारत (Hindi/Devanagari)
        s === 'in' || s.indexOf('india') === 0) {
        return 'India';
    }
    // Common misspellings / state names entered as country
    var indianStates = ['west bengal','bengal','kolkata','bihar','jharkhand','assam','odisha','uttar pradesh',
        'maharashtra','delhi','tripura','madhya pradesh','rajasthan','gujarat','kerala','punjab',
        'haryana','himachal','chhattisgarh','karnataka','jabalpur','guwahati'];
    for (var i = 0; i < indianStates.length; i++) {
        if (s.indexOf(indianStates[i]) !== -1) return 'India';
    }

    // Bangladesh
    if (s.indexOf('bangladesh') !== -1 || s.indexOf('dhaka') !== -1) return 'Bangladesh';
    // Nepal
    if (s.indexOf('nepal') !== -1) return 'Nepal';
    // Pakistan
    if (s.indexOf('pakistan') !== -1) return 'Pakistan';
    // USA
    if (s.indexOf('united states') !== -1 || s.indexOf('usa') !== -1 || s === 'us') return 'USA';
    // UK
    if (s.indexOf('united kingdom') !== -1 || s === 'uk') return 'UK';
    // UAE
    if (s.indexOf('uae') !== -1 || s.indexOf('emirates') !== -1 || s.indexOf('dubai') !== -1) return 'UAE';
    // Canada
    if (s.indexOf('canada') !== -1) return 'Canada';
    // Australia
    if (s.indexOf('australia') !== -1) return 'Australia';

    // Capitalise first letter for anything else
    return raw.trim().charAt(0).toUpperCase() + raw.trim().slice(1).toLowerCase();
}

function normalizeState(raw) {
    if (!raw) return null;
    var s = raw.trim().toLowerCase();
    if (!s) return null;

    var map = {
        'wb': 'West Bengal', 'west bengal': 'West Bengal', 'bengal': 'West Bengal',
        'jharkhand': 'Jharkhand', 'jharkhand ': 'Jharkhand',
        'bihar': 'Bihar', 'bihar ': 'Bihar',
        'uttar pradesh': 'Uttar Pradesh', 'up': 'Uttar Pradesh',
        'assam': 'Assam',
        'odisha': 'Odisha', 'orissa': 'Odisha',
        'tripura': 'Tripura',
        'maharashtra': 'Maharashtra',
        'delhi': 'Delhi', 'new delhi': 'Delhi',
        'gujarat': 'Gujarat',
        'rajasthan': 'Rajasthan',
        'madhya pradesh': 'Madhya Pradesh', 'mp': 'Madhya Pradesh',
        'kerala': 'Kerala',
        'karnataka': 'Karnataka',
        'punjab': 'Punjab',
        'haryana': 'Haryana',
        'himachal pradesh': 'Himachal Pradesh', 'himachal': 'Himachal Pradesh',
        'chhattisgarh': 'Chhattisgarh', 'chhatisgarh': 'Chhattisgarh',
        'andhra pradesh': 'Andhra Pradesh', 'ap': 'Andhra Pradesh',
        'telangana': 'Telangana',
        'goa': 'Goa',
        'manipur': 'Manipur',
        'meghalaya': 'Meghalaya',
        'nagaland': 'Nagaland',
        'mizoram': 'Mizoram',
        'sikkim': 'Sikkim',
        'arunachal pradesh': 'Arunachal Pradesh'
    };

    // Try exact lowercase match
    if (map[s]) return map[s];

    // Try partial match
    for (var key in map) {
        if (s.indexOf(key) !== -1 || key.indexOf(s) !== -1) return map[key];
    }

    // Capitalise as-is
    return raw.trim().charAt(0).toUpperCase() + raw.trim().slice(1);
}

// ============ FILE PROCESSING ============
function processFile(file) {
    var overlay = document.getElementById('loading-overlay');
    overlay.classList.remove('hidden');
    document.getElementById('data-status-text').innerText = 'Reading file...';

    var reader = new FileReader();

    reader.onerror = function() {
        overlay.classList.add('hidden');
        alert('Could not read the file. Please try again.');
    };

    reader.onload = function(e) {
        document.getElementById('data-status-text').innerText = 'Parsing spreadsheet...';

        // Give the UI a moment to paint the loading spinner before locking the thread
        setTimeout(function() {
            try {
                var data = new Uint8Array(e.target.result);
                var workbook = XLSX.read(data, { type: 'array', cellDates: true });

                // Try to find the right sheet
                var sheetName = workbook.SheetNames[0];
                for (var i = 0; i < workbook.SheetNames.length; i++) {
                    if (workbook.SheetNames[i].toLowerCase().indexOf('response') !== -1 ||
                        workbook.SheetNames[i].toLowerCase().indexOf('2025') !== -1) {
                        sheetName = workbook.SheetNames[i];
                        break;
                    }
                }

                var worksheet = workbook.Sheets[sheetName];
                globalRawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

                if (globalRawData.length < 2) {
                    overlay.classList.add('hidden');
                    alert('The spreadsheet seems empty. Found ' + globalRawData.length + ' rows in sheet "' + sheetName + '".');
                    return;
                }

                // Auto-detect column indices from header row
                detectColumns(globalRawData[0]);

                document.getElementById('timeframe-wrapper').classList.remove('hidden');
                recalculateData();

            } catch (error) {
                overlay.classList.add('hidden');
                alert('Error parsing spreadsheet:\n' + error.message + '\n\nPlease make sure this is a valid .xlsx or .csv file.');
                console.error('Full error:', error);
            }
        }, 50);
    };

    reader.readAsArrayBuffer(file);
}

// ============ LIVE SYNC PROCESSING ============
function fetchGoogleSheet(rawUrl) {
    if (!window.Papa) {
        alert('PapaParse library is missing. Cannot fetch CSV.');
        return;
    }

    var url = rawUrl;
    // Auto-correct standard edit links into export links
    if (url.indexOf('/edit') !== -1 && url.indexOf('docs.google.com/spreadsheets') !== -1) {
        url = url.substring(0, url.indexOf('/edit')) + '/export?format=csv';
    }

    var overlay = document.getElementById('loading-overlay');
    overlay.classList.remove('hidden');
    document.getElementById('data-status-text').innerText = 'Downloading live data...';

    // Helper to attempt fetch, with fallback proxy for file:// CORS blockers
    function performFetch(fetchUrl, useProxyOnFail) {
        
        // Aggressive Cache-Busting: Append a random timestamp to force Google/Browsers to skip cache
        var cbString = 'cb=' + new Date().getTime();
        var bustedUrl = fetchUrl + (fetchUrl.indexOf('?') !== -1 ? '&' : '?') + cbString;

        fetch(bustedUrl, { cache: 'no-store' })
            .then(function(response) {
                if (!response.ok) throw new Error("HTTP error " + response.status);
                return response.text();
            })
            .then(function(csvText) {
                // Google intercepts with an HTML login/error page if the sheet is private
                if (csvText.trim().toLowerCase().indexOf('<!doctype html>') === 0 || csvText.indexOf('<html') !== -1) {
                    throw new Error("Google returned a webpage instead of data. The sheet is likely set to 'Restricted' or requires organization login. Please make sure the link is fully Public.");
                }

                document.getElementById('data-status-text').innerText = 'Parsing live data...';
                
                Papa.parse(csvText, {
                    header: false,
                    skipEmptyLines: true,
                    complete: function(results) {
                        setTimeout(function() {
                            try {
                                globalRawData = results.data;
                                if (globalRawData.length < 2) {
                                    overlay.classList.add('hidden');
                                    alert('The live Google Sheet seems empty.');
                                    return;
                                }

                                detectColumns(globalRawData[0]);
                                document.getElementById('timeframe-wrapper').classList.remove('hidden');
                                recalculateData();
                            } catch(e) {
                                overlay.classList.add('hidden');
                                alert('Error evaluating live data:\n' + e.message);
                                console.error('Full Error:', e);
                            }
                        }, 50);
                    }
                });
            })
            .catch(function(err) {
                // If it failed and we haven't tried the proxy yet, try bypassing CORS!
                if (useProxyOnFail) {
                    console.warn("Direct fetch blocked by browser CORS. Attempting proxy bypass...", err);
                    var proxyUrl = 'https://api.allorigins.win/raw?url=' + encodeURIComponent(url);
                    performFetch(proxyUrl, false); // Try again with proxy, but don't loop infinitely
                } else {
                    overlay.classList.add('hidden');
                    // Check if it's running locally which breaks things permanently without a working proxy
                    if (window.location.protocol === 'file:') {
                        alert("Error downloading live data.\n\nBecause you opened this dashboard directly from your computer (file://), Google is blocking the connection for security.\n\nPlease drop your downloaded Excel file instead, or run the local server.");
                    } else {
                        var msg = err.message || 'Unknown network error / CORS blocker.';
                        alert('Error downloading from Google Sheets.\n\nDetails: ' + msg);
                    }
                    console.error("Fetch Live Error:", err);
                }
            });
    }

    // Start with direct fetch. If the origin is 'file:', it will naturally fail and fall back to the proxy.
    performFetch(url, true);
}

// ============ COLUMN AUTO-DETECTION ============
var COL = {
    timestamp: 0,
    gender: 6,
    country: 8,
    state: 9,
    isNewPatient: 14,
    source: 18
};

function detectColumns(headerRow) {
    // Try to auto-detect columns from headers
    for (var i = 0; i < headerRow.length; i++) {
        var h = String(headerRow[i] || '').toLowerCase();
        if (h.indexOf('timestamp') !== -1) COL.timestamp = i;
        if (h.indexOf('gender') !== -1) COL.gender = i;
        if (h.indexOf('country') !== -1) COL.country = i;
        if (h.indexOf('state') !== -1) COL.state = i;
        if (h.indexOf('new patient') !== -1) COL.isNewPatient = i;
        if (h.indexOf('hear about') !== -1 || h.indexOf('how did') !== -1) COL.source = i;
    }
}

// ============ DATA CALCULATION (with filter) ============
function recalculateData() {
    var filterVal = document.getElementById('timeframe-filter').value;
    dashboardData = initDashboardObject();
    var rows = globalRawData;
    var now = new Date();
    var validRows = 0;

    for (var i = 1; i < rows.length; i++) {
        var row = rows[i];
        if (!row || row.length < 3) continue;

        // Parse Date
        var rawDate = row[COL.timestamp];
        var rowDate = null;

        if (rawDate instanceof Date) {
            rowDate = rawDate;
        } else if (rawDate) {
            var dateStr = String(rawDate);
            // Try parsing "M/D/YYYY H:M:S" format
            var parts = dateStr.split(' ')[0].split('/');
            if (parts.length >= 3) {
                var m = parseInt(parts[0], 10) - 1;
                var d = parseInt(parts[1], 10);
                var y = parseInt(parts[2], 10);
                if (!isNaN(m) && !isNaN(d) && !isNaN(y) && y > 2000) {
                    rowDate = new Date(y, m, d);
                }
            }
            if (!rowDate || isNaN(rowDate.getTime())) {
                // Fallback: try native parse
                rowDate = new Date(dateStr);
                if (isNaN(rowDate.getTime())) rowDate = null;
            }
        }

        // Time Filter
        if (filterVal !== 'all' && rowDate) {
            if (filterVal === 'last30') {
                var diffDays = (now - rowDate) / (1000 * 60 * 60 * 24);
                if (diffDays > 30 || diffDays < 0) continue;
            } else {
                var fltY = parseInt(filterVal, 10);
                if (!isNaN(fltY) && rowDate.getFullYear() !== fltY) continue;
            }
        }

        validRows++;
        dashboardData.totalPatients++;

        // Extract fields
        var genderRaw = String(row[COL.gender] || '').split('/')[0].trim();
        var countryRaw = String(row[COL.country] || '').trim();
        var stateRaw = String(row[COL.state] || '').trim();
        var isNewStr = String(row[COL.isNewPatient] || '').toLowerCase();
        var sourceStr = String(row[COL.source] || '');

        // Normalize gender (take first word before /)
        var gender = genderRaw.split(' ')[0];
        if (gender === 'Male' || gender === 'Female' || gender === 'Other') {
            dashboardData.demographics.gender[gender] = (dashboardData.demographics.gender[gender] || 0) + 1;
        } else if (gender) {
            // Try to normalize common variations
            var gl = gender.toLowerCase();
            if (gl === 'male' || gl === 'm') gender = 'Male';
            else if (gl === 'female' || gl === 'f') gender = 'Female';
            else gender = 'Other';
            dashboardData.demographics.gender[gender] = (dashboardData.demographics.gender[gender] || 0) + 1;
        }

        var country = normalizeCountry(countryRaw);
        if (country) dashboardData.demographics.country[country] = (dashboardData.demographics.country[country] || 0) + 1;

        var state = normalizeState(stateRaw);
        if (state) dashboardData.demographics.state[state] = (dashboardData.demographics.state[state] || 0) + 1;

        // Acquisition Source
        var broadBucket = classifyBroadSource(sourceStr);
        dashboardData.broadSource[broadBucket] = (dashboardData.broadSource[broadBucket] || 0) + 1;

        var isNew = isNewStr.indexOf('yes') !== -1;

        if (broadBucket === 'Social Media') {
            dashboardData.socialPatients++;
            
            if (state) dashboardData.socialStates[state] = (dashboardData.socialStates[state] || 0) + 1;
            if (rowDate) {
                var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
                var dName = days[rowDate.getDay()];
                dashboardData.socialDays[dName] = (dashboardData.socialDays[dName] || 0) + 1;
            }

            var platform = classifySocialPlatform(sourceStr);
            // Only add if it's one of the 5 known platforms (null = generic, skip)
            if (platform !== null) {
                dashboardData.socialPlatform[platform] = (dashboardData.socialPlatform[platform] || 0) + 1;

                if (!dashboardData.platformConversion[platform]) dashboardData.platformConversion[platform] = { new: 0, old: 0 };
                if (isNew) dashboardData.platformConversion[platform].new++;
                else dashboardData.platformConversion[platform].old++;

                if (gender) {
                    if (!dashboardData.platformGender[platform]) dashboardData.platformGender[platform] = {};
                    dashboardData.platformGender[platform][gender] = (dashboardData.platformGender[platform][gender] || 0) + 1;
                }

                if (rowDate && !isNaN(rowDate.getTime())) {
                    var pym = rowDate.getFullYear() + '-' + String(rowDate.getMonth() + 1).padStart(2, '0');
                    if (!dashboardData.platformTimeline[pym]) dashboardData.platformTimeline[pym] = {};
                    dashboardData.platformTimeline[pym][platform] = (dashboardData.platformTimeline[pym][platform] || 0) + 1;
                }
            }

            if (isNew) {
                dashboardData.socialNewPatients++;
                dashboardData.socialNewVsOld.new++;
            } else {
                dashboardData.socialNewVsOld.old++;
            }
        }

        // Timeline
        if (rowDate && !isNaN(rowDate.getTime())) {
            var ym = rowDate.getFullYear() + '-' + String(rowDate.getMonth() + 1).padStart(2, '0');
            if (!dashboardData.timeline[ym]) dashboardData.timeline[ym] = { total: 0, social: 0 };
            if (!dashboardData.socialConversionTimeline[ym]) dashboardData.socialConversionTimeline[ym] = { new: 0, old: 0 };
            
            dashboardData.timeline[ym].total++;
            if (broadBucket === 'Social Media') {
                dashboardData.timeline[ym].social++;
                if (isNew) {
                    dashboardData.socialConversionTimeline[ym].new++;
                } else {
                    dashboardData.socialConversionTimeline[ym].old++;
                }
            }
        }
    }

    if (validRows > 0) {
        updateUI();
    } else {
        document.getElementById('loading-overlay').classList.add('hidden');
        alert('No records found for the selected timeframe. Try "All Time".');
    }
}

// ============ ANIMATED COUNT-UP ============
function animateCount(el, target, duration) {
    var startTime = null;
    var easeOut = function(t) { return 1 - Math.pow(1 - t, 3); };
    function step(ts) {
        if (!startTime) startTime = ts;
        var p = Math.min((ts - startTime) / duration, 1);
        el.innerText = Math.round(easeOut(p) * target).toLocaleString();
        if (p < 1) requestAnimationFrame(step);
    }
    requestAnimationFrame(step);
}

// ============ UI UPDATE ============
function updateUI() {
    // Stats — animated
    animateCount(document.getElementById('stat-total-patients'), dashboardData.totalPatients, 1200);
    animateCount(document.getElementById('stat-social-total'), dashboardData.socialPatients, 1200);
    animateCount(document.getElementById('stat-social-new'), dashboardData.socialNewPatients, 1200);

    var socialPct = dashboardData.totalPatients > 0
        ? ((dashboardData.socialPatients / dashboardData.totalPatients) * 100).toFixed(1)
        : 0;
    document.getElementById('stat-social-percentage').innerText = socialPct + '%';

    var newPct = dashboardData.socialPatients > 0
        ? ((dashboardData.socialNewPatients / dashboardData.socialPatients) * 100).toFixed(1)
        : 0;
    document.getElementById('stat-social-new-pct').innerText = newPct + '%';

    // New Text Metrics
    var sortedStates = sortEntries(dashboardData.socialStates);
    document.getElementById('stat-top-state').innerText = sortedStates.length > 0 ? sortedStates[0][0] : '-';

    var sortedDays = sortEntries(dashboardData.socialDays);
    document.getElementById('stat-peak-day').innerText = sortedDays.length > 0 && sortedDays[0][1] > 0 ? sortedDays[0][0] : '-';

    // Platform Insights Data
    var sortedPlatforms = sortEntries(dashboardData.socialPlatform);
    document.getElementById('stat-champion-platform').innerText = sortedPlatforms.length > 0 ? sortedPlatforms[0][0] : '-';

    var highestRetainer = '-';
    var highestOld = -1;
    for (var p in dashboardData.platformConversion) {
        if (dashboardData.platformConversion[p].old > highestOld) {
            highestOld = dashboardData.platformConversion[p].old;
            highestRetainer = p;
        }
    }
    document.getElementById('stat-platform-retainer').innerText = highestOld > 0 ? highestRetainer : '-';

    // Charts
    renderOverviewCharts();
    renderDemographicsCharts();
    renderPlatformCharts();

    // Show Dashboard
    document.getElementById('upload-splash').style.display = 'none';
    var body = document.getElementById('dashboard-body');
    body.classList.remove('glass-hidden');
    body.style.display = '';

    document.getElementById('loading-overlay').classList.add('hidden');

    // Show PDF export and Footer
    var pdfBtn = document.getElementById('pdf-export-btn');
    if(pdfBtn) pdfBtn.classList.remove('hidden');
    
    var footer = document.getElementById('dashboard-footer');
    if(footer) {
        footer.classList.remove('hidden');
        var d = new Date();
        document.getElementById('footer-date').innerText = d.toLocaleDateString() + ' ' + d.toLocaleTimeString();
    }

    var indicator = document.querySelector('.status-indicator');
    indicator.className = 'status-indicator connected';
    document.getElementById('data-status-text').innerText = dashboardData.totalPatients.toLocaleString() + ' records loaded';
}

// ============ EXPORT PDF ============
async function exportPDF() {
    var btn = document.getElementById('pdf-export-btn');
    var originalText = btn.innerHTML;
    
    var dateStr = new Date().toISOString().split('T')[0];
    var suggestedName = 'PBHRC_Social_Media_Report_' + dateStr + '.pdf';
    var fileHandle = null;

    // 1. We must ask the user where to save FIRST (browsers require this happen immediately after click)
    try {
        if (window.showSaveFilePicker) {
            fileHandle = await window.showSaveFilePicker({
                suggestedName: suggestedName,
                types: [{
                    description: 'PDF Document',
                    accept: {'application/pdf': ['.pdf']}
                }]
            });
        }
    } catch (e) {
        // If user hits "Cancel" on the save prompt, just stop
        if (e.name === 'AbortError') return; 
        console.warn('Native save API not supported/failed, will use fallback.', e);
    }

    btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Generating...';
    btn.disabled = true;

    // Temporarily hide the PDF button so it doesn't show in the PDF
    btn.style.display = 'none';

    try {
        var { jsPDF } = window.jspdf;
        var element = document.querySelector('.main-content');

        // Ensure we capture from the top
        element.scrollTop = 0;

        var canvas = await html2canvas(element, {
            scale: 1.5, 
            useCORS: true,
            backgroundColor: '#0f172a',
            scrollY: 0
        });

        // Use JPEG to keep size tiny
        var imgData = canvas.toDataURL('image/jpeg', 0.85);
        
        var pdf = new jsPDF('l', 'mm', 'a3'); 
        var pdfWidth = pdf.internal.pageSize.getWidth();
        var pdfHeight = (canvas.height * pdfWidth) / canvas.width;

        pdf.addImage(imgData, 'JPEG', 0, 0, pdfWidth, pdfHeight);
        
        if (fileHandle) {
            // 2. Use Native Windows Saving API (guarantees .pdf extension)
            var blob = pdf.output('blob');
            var writable = await fileHandle.createWritable();
            await writable.write(blob);
            await writable.close();
        } else {
            // Fallback for older browsers
            pdf.save(suggestedName);
        }

    } catch (err) {
        console.error('PDF Error:', err);
        alert('Could not generate PDF. Please check the console for details.');
    } finally {
        btn.style.display = '';
        btn.innerHTML = originalText;
        btn.disabled = false;
    }
}

// ============ CHART RENDERING ============
function getChartLabelColor() {
    return document.documentElement.getAttribute('data-theme') === 'light' ? '#4a6175' : '#f0f4f8';
}
function getGridColor() {
    return document.documentElement.getAttribute('data-theme') === 'light' ? 'rgba(0,0,0,0.06)' : 'rgba(255,255,255,0.05)';
}
var chartColors = ['#176eb5', '#2ecc71', '#f39c12', '#E63946', '#9b59b6', '#e84393', '#1abc9c', '#7f8c8d'];
var platformColors = {
    'Facebook': '#4267B2',
    'Instagram': '#C13584',
    'Facebook & Instagram': '#8B5CF6',
    'YouTube': '#FF0000',
    'WhatsApp': '#25D366',
    'Generic Social Media': '#f39c12'
};

function renderOverviewCharts() {
    // 1. Broad Source Doughnut
    if (stateCharts.pieBroad) stateCharts.pieBroad.destroy();
    var sortedB = sortEntries(dashboardData.broadSource);
    stateCharts.pieBroad = new Chart(document.getElementById('sourcePieChart').getContext('2d'), {
        type: 'doughnut',
        data: {
            labels: sortedB.map(function(x) { return x[0]; }),
            datasets: [{
                data: sortedB.map(function(x) { return x[1]; }),
                backgroundColor: chartColors,
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '65%',
            plugins: { legend: { position: 'right', labels: { color: getChartLabelColor(), padding: 12 } } }
        }
    });

    // 2. Social Platform Pie
    if (stateCharts.piePlatform) stateCharts.piePlatform.destroy();
    var sortedP = sortEntries(dashboardData.socialPlatform);
    var pColors = sortedP.map(function(x) { return platformColors[x[0]] || '#7f8c8d'; });
    stateCharts.piePlatform = new Chart(document.getElementById('platformPieChart').getContext('2d'), {
        type: 'pie',
        data: {
            labels: sortedP.map(function(x) { return x[0]; }),
            datasets: [{
                data: sortedP.map(function(x) { return x[1]; }),
                backgroundColor: pColors,
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'right', labels: { color: getChartLabelColor(), padding: 12 } } }
        }
    });

    // 3. New vs Old Bar
    if (stateCharts.barNewOld) stateCharts.barNewOld.destroy();
    stateCharts.barNewOld = new Chart(document.getElementById('newVSoldChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: ['From Social Media'],
            datasets: [
                { label: 'New Patients', data: [dashboardData.socialNewVsOld.new], backgroundColor: '#176eb5', borderRadius: 8 },
                { label: 'Existing', data: [dashboardData.socialNewVsOld.old], backgroundColor: '#E63946', borderRadius: 8 }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: getChartLabelColor() } } },
            scales: {
                y: { grid: { color: getGridColor() } },
                x: { grid: { display: false } }
            }
        }
    });

    // 4. Timeline
    if (stateCharts.timeline) stateCharts.timeline.destroy();
    var timelineKeys = Object.keys(dashboardData.timeline).sort();
    var tLabels = timelineKeys.map(function(k) {
        var parts = k.split('-');
        var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return months[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
    });
    stateCharts.timeline = new Chart(document.getElementById('timelineChart').getContext('2d'), {
        type: 'line',
        data: {
            labels: tLabels,
            datasets: [
                {
                    label: 'Total Registrations',
                    data: timelineKeys.map(function(k) { return dashboardData.timeline[k].total; }),
                    borderColor: 'rgba(255,255,255,0.4)',
                    backgroundColor: getGridColor(),
                    fill: true,
                    tension: 0.4,
                    borderWidth: 2
                },
                {
                    label: 'Social Media',
                    data: timelineKeys.map(function(k) { return dashboardData.timeline[k].social; }),
                    borderColor: '#176eb5',
                    backgroundColor: 'rgba(23,110,181,0.2)',
                    fill: true,
                    tension: 0.4,
                    borderWidth: 3,
                    pointBackgroundColor: '#176eb5',
                    pointRadius: 3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { labels: { color: getChartLabelColor() } } },
            scales: {
                y: { grid: { color: getGridColor() } },
                x: { grid: { color: getGridColor() } }
            }
        }
    });

    // 5. Social Conversion Timeline
    if (stateCharts.socialConversionTimeline) stateCharts.socialConversionTimeline.destroy();
    var convKeys = Object.keys(dashboardData.socialConversionTimeline).sort();
    var cLabels = convKeys.map(function(k) {
        var parts = k.split('-');
        var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
        return months[parseInt(parts[1], 10) - 1] + ' ' + parts[0];
    });
    
    stateCharts.socialConversionTimeline = new Chart(document.getElementById('socialConversionTimelineChart').getContext('2d'), {
        type: 'line',
        data: {
            labels: cLabels,
            datasets: [
                {
                    label: 'New Patients (Social Media)',
                    data: convKeys.map(function(k) { return dashboardData.socialConversionTimeline[k].new; }),
                    borderColor: '#176eb5',
                    backgroundColor: 'rgba(23,110,181,0.2)',
                    fill: true,
                    tension: 0.4,
                    borderWidth: 3,
                    pointBackgroundColor: '#176eb5',
                    pointRadius: 3
                },
                {
                    label: 'Existing Patients (Social Media)',
                    data: convKeys.map(function(k) { return dashboardData.socialConversionTimeline[k].old; }),
                    borderColor: '#E63946',
                    backgroundColor: 'transparent',
                    tension: 0.4,
                    borderWidth: 3,
                    pointBackgroundColor: '#E63946',
                    pointRadius: 3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: { legend: { labels: { color: getChartLabelColor() } } },
            scales: {
                y: { grid: { color: getGridColor() }, beginAtZero: true },
                x: { grid: { color: getGridColor() } }
            }
        }
    });
}

function renderDemographicsCharts() {
    // Gender
    if (stateCharts.genderPie) stateCharts.genderPie.destroy();
    var sortedG = sortEntries(dashboardData.demographics.gender);
    stateCharts.genderPie = new Chart(document.getElementById('genderPieChart').getContext('2d'), {
        type: 'doughnut',
        data: {
            labels: sortedG.map(function(x) { return x[0]; }),
            datasets: [{ data: sortedG.map(function(x) { return x[1]; }), backgroundColor: ['#E63946', '#176eb5', '#f39c12'], borderWidth: 0 }]
        },
        options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'right', labels: { color: getChartLabelColor() } } } }
    });

    // Gender by Platform
    if (stateCharts.genderPlatformPie) stateCharts.genderPlatformPie.destroy();
    
    var gpLabels = [];
    var gpDataMale = [];
    var gpDataFemale = [];
    var platformsWithGender = Object.keys(dashboardData.platformGender);
    for(var i=0; i<platformsWithGender.length; i++) {
        var plat = platformsWithGender[i];
        gpLabels.push(plat);
        gpDataMale.push(dashboardData.platformGender[plat]['Male'] || 0);
        gpDataFemale.push(dashboardData.platformGender[plat]['Female'] || 0);
    }
    stateCharts.genderPlatformPie = new Chart(document.getElementById('genderPlatformPieChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: gpLabels,
            datasets: [
                { label: 'Male', data: gpDataMale, backgroundColor: '#176eb5' },
                { label: 'Female', data: gpDataFemale, backgroundColor: '#f39c12' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: { stacked: true, ticks: { color: getChartLabelColor() }, grid: { display: false } },
                y: { stacked: true, ticks: { color: getChartLabelColor() }, grid: { color: getGridColor() } }
            },
            plugins: {
                legend: { position: 'bottom', labels: { color: getChartLabelColor() } }
            }
        }
    });

    // Country Top 10
    if (stateCharts.countryBar) stateCharts.countryBar.destroy();
    var sortedC = sortEntries(dashboardData.demographics.country).slice(0, 10);
    stateCharts.countryBar = new Chart(document.getElementById('countryBarChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: sortedC.map(function(x) { return x[0]; }),
            datasets: [{ label: 'Patients', data: sortedC.map(function(x) { return x[1]; }), backgroundColor: '#2ecc71', borderRadius: 4 }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            indexAxis: 'y',
            plugins: { legend: { display: false } },
            scales: { x: { grid: { color: getGridColor() } } }
        }
    });

    // State Breakdown Top 15
    if (stateCharts.stateBar) stateCharts.stateBar.destroy();
    var sortedS = sortEntries(dashboardData.demographics.state).slice(0, 15);
    stateCharts.stateBar = new Chart(document.getElementById('stateBarChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: sortedS.map(function(x) { return x[0]; }),
            datasets: [{ label: 'Patients', data: sortedS.map(function(x) { return x[1]; }), backgroundColor: '#176eb5', borderRadius: 4 }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { grid: { color: getGridColor() } },
                x: { grid: { display: false } }
            }
        }
    });
}

function renderPlatformCharts() {
    var tc = getChartLabelColor();
    var gc = getGridColor();

    if (stateCharts.platformConversionLine) stateCharts.platformConversionLine.destroy();
    if (stateCharts.platformDistributionBar) stateCharts.platformDistributionBar.destroy();

    var platforms = Object.keys(dashboardData.platformConversion);
    var pConvNew = platforms.map(function(p) { return dashboardData.platformConversion[p].new; });
    var pConvOld = platforms.map(function(p) { return dashboardData.platformConversion[p].old; });

    stateCharts.platformConversionLine = new Chart(document.getElementById('platformConversionChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: platforms,
            datasets: [
                { label: 'New Patients', data: pConvNew, backgroundColor: '#176eb5', borderRadius: 4 },
                { label: 'Existing Patients', data: pConvOld, backgroundColor: '#f39c12', borderRadius: 4 }
            ]
        },
        options: {
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: tc }, grid: { display: false } },
                y: { ticks: { color: tc }, grid: { color: gc } }
            },
            plugins: { legend: { labels: { color: tc } } }
        }
    });

    var timeKeys = Object.keys(dashboardData.platformTimeline).sort();
    var trackedPlatforms = ['Facebook', 'Instagram', 'Facebook & Instagram', 'YouTube', 'WhatsApp', 'LinkedIn'];
    
    var datasets = trackedPlatforms.map(function(plat) {
        return {
            label: plat,
            data: timeKeys.map(function(tk) {
                return dashboardData.platformTimeline[tk][plat] || 0;
            }),
            backgroundColor: platformColors[plat] || '#ccc',
            fill: true,
            tension: 0.3
        };
    });

    stateCharts.platformDistributionBar = new Chart(document.getElementById('platformDistributionChart').getContext('2d'), {
        type: 'line',
        data: { labels: timeKeys, datasets: datasets },
        options: {
            maintainAspectRatio: false,
            scales: {
                x: { ticks: { color: tc }, grid: { display: false } },
                y: { stacked: true, ticks: { color: tc, stepSize: 10 }, grid: { color: gc } }
            },
            plugins: {
                legend: { position: 'bottom', labels: { color: tc } }
            }
        }
    });
}

// Helper: sort object entries by value descending
function sortEntries(obj) {
    var entries = [];
    for (var key in obj) {
        if (obj.hasOwnProperty(key)) entries.push([key, obj[key]]);
    }
    entries.sort(function(a, b) { return b[1] - a[1]; });
    return entries;
}

// String.prototype.padStart polyfill for older browsers
if (!String.prototype.padStart) {
    String.prototype.padStart = function(targetLength, padString) {
        targetLength = targetLength >> 0;
        padString = String(typeof padString !== 'undefined' ? padString : ' ');
        if (this.length >= targetLength) return String(this);
        targetLength = targetLength - this.length;
        if (targetLength > padString.length) padString += padString.repeat(targetLength / padString.length);
        return padString.slice(0, targetLength) + String(this);
    };
}
