<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>Spare & Tractor Service Manager</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- SheetJS for Excel Generation -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Inter Font -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    
    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#f0f9ff',
                            100: '#e0f2fe',
                            500: '#0284c7',
                            600: '#0369a1',
                            700: '#075985',
                        },
                        amberTheme: {
                            50: '#fffbebe6',
                            100: '#fef3c7',
                            500: '#f59e0b',
                            600: '#d97706',
                            700: '#b45309',
                        }
                    },
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    }
                }
            }
        }
    </script>
    <style>
        /* Custom Mobile Fixes */
        body {
            -webkit-tap-highlight-color: transparent;
            font-family: 'Inter', sans-serif;
        }
        .pb-safe {
            padding-bottom: env(safe-area-inset-bottom, 1rem);
        }
    </style>
</head>
<body class="bg-slate-50 dark:bg-slate-900 text-slate-800 dark:text-slate-100 min-h-screen flex flex-col antialiased selection:bg-brand-500 selection:text-white">

    <!-- TOP NAVBAR -->
    <header class="sticky top-0 z-30 bg-white/90 dark:bg-slate-800/90 backdrop-blur-md border-b border-slate-200 dark:border-slate-700 px-4 py-3 flex items-center justify-between shadow-sm">
        <div class="flex items-center space-x-3">
            <div class="bg-brand-600 text-white p-2 rounded-xl shadow-md">
                <i class="fa-solid font-bold fa-tractor text-xl"></i>
            </div>
            <div>
                <h1 class="font-bold text-lg leading-tight text-slate-900 dark:text-white">SpareManager Pro</h1>
                <p class="text-xs text-slate-500 dark:text-slate-400">Live Sync & Excel Manager</p>
            </div>
        </div>

        <div class="flex items-center space-x-2">
            <!-- Sync Status Indicator -->
            <span id="sync-badge" class="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300">
                <span class="w-2 h-2 mr-1.5 bg-emerald-500 rounded-full animate-pulse"></span> Ready
            </span>
            <!-- Dark Mode Toggle -->
            <button id="theme-toggle" class="p-2 text-slate-500 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-700 rounded-xl">
                <i class="fa-solid fa-moon text-lg dark:hidden"></i>
                <i class="fa-solid fa-sun text-lg hidden dark:block"></i>
            </button>
        </div>
    </header>

    <!-- MAIN WRAPPER -->
    <div class="flex-1 flex max-w-7xl w-full mx-auto pb-24 md:pb-6">

        <!-- DESKTOP SIDEBAR -->
        <aside class="hidden md:flex flex-col w-64 bg-white dark:bg-slate-800 border-r border-slate-200 dark:border-slate-700 p-4 space-y-1">
            <p class="text-xs font-semibold text-slate-400 uppercase tracking-wider px-3 mb-2">Navigation</p>
            
            <button onclick="switchTab('panel1')" id="nav-desk-panel1" class="nav-desk-btn flex items-center space-x-3 w-full px-3 py-2.5 rounded-xl font-medium text-sm transition-all text-brand-600 bg-brand-50 dark:bg-brand-900/30 dark:text-brand-400">
                <i class="fa-solid fa-wrench w-5 text-center"></i>
                <span>1. Spare Repair Service</span>
            </button>

            <button onclick="switchTab('panel2')" id="nav-desk-panel2" class="nav-desk-btn flex items-center space-x-3 w-full px-3 py-2.5 rounded-xl font-medium text-sm text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-all">
                <i class="fa-solid fa-boxes-stacked w-5 text-center"></i>
                <span>2. Inventory & Fitting</span>
            </button>

            <button onclick="switchTab('panel3')" id="nav-desk-panel3" class="nav-desk-btn flex items-center space-x-3 w-full px-3 py-2.5 rounded-xl font-medium text-sm text-amber-600 dark:text-amber-400 hover:bg-amber-50 dark:hover:bg-amber-900/20 transition-all">
                <i class="fa-solid fa-bell w-5 text-center"></i>
                <span>3. Service Tracker</span>
            </button>

            <button onclick="switchTab('panel4')" id="nav-desk-panel4" class="nav-desk-btn flex items-center space-x-3 w-full px-3 py-2.5 rounded-xl font-medium text-sm text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-all">
                <i class="fa-solid fa-table-list w-5 text-center"></i>
                <span>4. View Data & Excel</span>
            </button>

            <div class="mt-auto pt-4 border-t border-slate-200 dark:border-slate-700">
                <div class="bg-slate-100 dark:bg-slate-700/50 p-3 rounded-xl text-xs space-y-1">
                    <p class="font-semibold text-slate-700 dark:text-slate-200"><i class="fa-solid fa-cloud text-brand-500 mr-1"></i> Sync Status</p>
                    <p class="text-slate-500 dark:text-slate-400">Live Cloud storage active for iPhone & PC.</p>
                </div>
            </div>
        </aside>

        <!-- CONTENT CONTAINER -->
        <main class="flex-1 p-4 md:p-6 overflow-y-auto max-w-4xl mx-auto">

            <!-- STATS QUICK BAR -->
            <div class="grid grid-cols-3 gap-3 mb-6">
                <div class="bg-white dark:bg-slate-800 p-3.5 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center justify-between">
                    <div>
                        <p class="text-xs text-slate-500 dark:text-slate-400">Repairs</p>
                        <p id="stat-repairs" class="text-xl font-bold text-slate-900 dark:text-white">0</p>
                    </div>
                    <div class="p-2.5 bg-blue-50 dark:bg-blue-900/30 text-blue-600 dark:text-blue-400 rounded-xl">
                        <i class="fa-solid fa-wrench"></i>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-3.5 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center justify-between">
                    <div>
                        <p class="text-xs text-slate-500 dark:text-slate-400">Stock Entries</p>
                        <p id="stat-inventory" class="text-xl font-bold text-slate-900 dark:text-white">0</p>
                    </div>
                    <div class="p-2.5 bg-emerald-50 dark:bg-emerald-900/30 text-emerald-600 dark:text-emerald-400 rounded-xl">
                        <i class="fa-solid fa-box"></i>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-3.5 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center justify-between">
                    <div>
                        <p class="text-xs text-slate-500 dark:text-slate-400">Due Services</p>
                        <p id="stat-services" class="text-xl font-bold text-amber-600 dark:text-amber-400">0</p>
                    </div>
                    <div class="p-2.5 bg-amber-50 dark:bg-amber-900/30 text-amber-600 dark:text-amber-400 rounded-xl">
                        <i class="fa-solid fa-clock-rotate-left"></i>
                    </div>
                </div>
            </div>

            <!-- PANEL 1: SPARE REPAIR ENTRY -->
            <section id="panel1" class="tab-content space-y-4">
                <div class="bg-white dark:bg-slate-800 p-5 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm">
                    <div class="flex items-center space-x-3 mb-4 pb-3 border-b border-slate-100 dark:border-slate-700">
                        <div class="p-2 bg-brand-500 text-white rounded-xl">
                            <i class="fa-solid fa-wrench"></i>
                        </div>
                        <div>
                            <h2 class="text-lg font-bold text-slate-900 dark:text-white">Spare Repair Service Entry</h2>
                            <p class="text-xs text-slate-500 dark:text-slate-400">Sheet 1: Party spare item repair intake log</p>
                        </div>
                    </div>

                    <form id="form-repair" onsubmit="handleRepairSubmit(event)" class="space-y-4">
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Party Name (Grahak Ka Naam) *</label>
                                <input type="text" id="r-party" required placeholder="e.g. Ramesh Kumar" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none">
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Tractor Chassis Number *</label>
                                <input type="text" id="r-chassis" required placeholder="e.g. CH123456" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none">
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Tractor Model *</label>
                                <input type="text" id="r-model" required placeholder="e.g. Mahindra 575 DI" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none">
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Repair Date *</label>
                                <input type="date" id="r-date" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none">
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Item Category *</label>
                                <select id="r-category" onchange="updateCompanyOptions('r-category', 'r-company')" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none">
                                    <option value="Battery">Battery</option>
                                    <option value="Self Starter">Self Starter</option>
                                    <option value="Alternator">Alternator</option>
                                </select>
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Company Name *</label>
                                <select id="r-company" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none">
                                    <!-- Dynamic Options -->
                                </select>
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Tractor Running Hours</label>
                                <input type="number" id="r-hours" placeholder="e.g. 1250" class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none">
                            </div>

                            <div class="md:col-span-2">
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Complaint / Problem Description *</label>
                                <textarea id="r-complaint" rows="2" required placeholder="Self not cranking / Battery discharge problem..." class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-brand-500 outline-none"></textarea>
                            </div>
                        </div>

                        <button type="submit" class="w-full bg-brand-600 hover:bg-brand-700 text-white font-semibold py-3 px-4 rounded-xl shadow-lg shadow-brand-500/20 active:scale-[0.99] transition-all flex items-center justify-center space-x-2">
                            <i class="fa-solid fa-floppy-disk"></i>
                            <span>Save Repair Record to Sheet 1</span>
                        </button>
                    </form>
                </div>
            </section>

            <!-- PANEL 2: INVENTORY & FITTING ENTRY -->
            <section id="panel2" class="tab-content hidden space-y-4">
                <div class="bg-white dark:bg-slate-800 p-5 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm">
                    <div class="flex items-center space-x-3 mb-4 pb-3 border-b border-slate-100 dark:border-slate-700">
                        <div class="p-2 bg-emerald-500 text-white rounded-xl">
                            <i class="fa-solid fa-boxes-stacked"></i>
                        </div>
                        <div>
                            <h2 class="text-lg font-bold text-slate-900 dark:text-white">Spare Inventory & Fitting Control</h2>
                            <p class="text-xs text-slate-500 dark:text-slate-400">Sheet 2: Spare Stock In & Fitted Out log</p>
                        </div>
                    </div>

                    <form id="form-inventory" onsubmit="handleInventorySubmit(event)" class="space-y-4">
                        <!-- Action Switcher -->
                        <div>
                            <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1.5">Inventory Action Type</label>
                            <div class="grid grid-cols-2 gap-3">
                                <label class="cursor-pointer border border-slate-300 dark:border-slate-600 rounded-xl p-3 flex items-center space-x-2 has-[:checked]:border-emerald-500 has-[:checked]:bg-emerald-50 dark:has-[:checked]:bg-emerald-900/20">
                                    <input type="radio" name="inv-type" value="Stock In" checked onchange="toggleInventoryFields()" class="accent-emerald-600">
                                    <div>
                                        <p class="text-xs font-bold text-slate-800 dark:text-slate-200">Stock Received (In)</p>
                                        <p class="text-[10px] text-slate-500">Add spare to store inventory</p>
                                    </div>
                                </label>

                                <label class="cursor-pointer border border-slate-300 dark:border-slate-600 rounded-xl p-3 flex items-center space-x-2 has-[:checked]:border-emerald-500 has-[:checked]:bg-emerald-50 dark:has-[:checked]:bg-emerald-900/20">
                                    <input type="radio" name="inv-type" value="Fitted/Out" onchange="toggleInventoryFields()" class="accent-emerald-600">
                                    <div>
                                        <p class="text-xs font-bold text-slate-800 dark:text-slate-200">Fitted in Tractor (Out)</p>
                                        <p class="text-[10px] text-slate-500">Attach spare to customer tractor</p>
                                    </div>
                                </label>
                            </div>
                        </div>

                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Item Category *</label>
                                <select id="i-category" onchange="updateCompanyOptions('i-category', 'i-company')" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-emerald-500 outline-none">
                                    <option value="Battery">Battery</option>
                                    <option value="Self Starter">Self Starter</option>
                                    <option value="Alternator">Alternator</option>
                                </select>
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Company Name *</label>
                                <select id="i-company" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-emerald-500 outline-none">
                                    <!-- Dynamic Options -->
                                </select>
                            </div>

                            <div>
                                <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Action Date *</label>
                                <input type="date" id="i-date" required class="w-full px-3.5 py-2.5 rounded-xl border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-emerald-500 outline-none">
                            </div>

                            <!-- Conditional Tractor Info (Out only) -->
                            <div id="inv-fitting-fields" class="hidden md:col-span-2 grid grid-cols-1 md:grid-cols-3 gap-4 bg-slate-50 dark:bg-slate-700/40 p-3 rounded-xl border border-slate-200 dark:border-slate-600">
                                <div>
                                    <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Chassis Number</label>
                                    <input type="text" id="i-chassis" placeholder="e.g. CH998877" class="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-xs outline-none">
                                </div>
                                <div>
                                    <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Tractor Model</label>
                                    <input type="text" id="i-model" placeholder="e.g. Swaraj 744 FE" class="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-xs outline-none">
                                </div>
                                <div>
                                    <label class="block text-xs font-semibold text-slate-600 dark:text-slate-300 mb-1">Party Name</label>
                                    <input type="text" id="i-party" placeholder="e.g. Suresh Patel" class="w-full px-3 py-2 rounded-lg border border-slate-300 dark:border-slate-600 bg-white dark:bg-slate-700 text-slate-900 dark:text-white text-xs outline-none">
                                </div>
                            </div>
                        </div>

                        <button type="submit" class="w-full bg-emerald-600 hover:bg-emerald-700 text-white font-semibold py-3 px-4 rounded-xl shadow-lg shadow-emerald-500/20 active:scale-[0.99] transition-all flex items-center justify-center space-x-2">
                            <i class="fa-solid fa-square-plus"></i>
                            <span>Save Inventory Entry to Sheet 2</span>
                        </button>
                    </form>
                </div>
            </section>

            <!-- PANEL 3: TRACTOR SERVICE REMINDER (DISTINCT AMBER THEME) -->
            <section id="panel3" class="tab-content hidden space-y-4">
                <div class="bg-amberTheme-50 dark:bg-amber-950/40 p-5 rounded-2xl border-2 border-amber-400 dark:border-amber-600 shadow-md">
                    <div class="flex items-center space-x-3 mb-4 pb-3 border-b border-amber-200 dark:border-amber-800/60">
                        <div class="p-2.5 bg-amber-600 text-white rounded-xl shadow-md">
                            <i class="fa-solid fa-bell-ring text-lg animate-bounce"></i>
                        </div>
                        <div>
                            <h2 class="text-lg font-bold text-amber-900 dark:text-amber-200">Tractor Service Due & Tracker</h2>
                            <p class="text-xs text-amber-700 dark:text-amber-400">Sheet 3: Scheduled customer tractor services & status</p>
                        </div>
                    </div>

                    <form id="form-service" onsubmit="handleServiceSubmit(event)" class="space-y-4">
                        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <div>
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Party Name (Customer) *</label>
                                <input type="text" id="s-party" required placeholder="e.g. Vikram Singh" class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none">
                            </div>

                            <div>
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Tractor Model *</label>
                                <input type="text" id="s-model" required placeholder="e.g. John Deere 5050D" class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none">
                            </div>

                            <div>
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Phone Number *</label>
                                <input type="tel" id="s-phone" required placeholder="e.g. 9876543210" class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none">
                            </div>

                            <div>
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Scheduled Service Due Date *</label>
                                <input type="date" id="s-duedate" required class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none">
                            </div>

                            <div class="md:col-span-2">
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Address / Location</label>
                                <input type="text" id="s-address" placeholder="e.g. Village Rampur, Near Main Tank" class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none">
                            </div>

                            <div class="md:col-span-2">
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Service Work / Complaint Details *</label>
                                <textarea id="s-complaint" rows="2" required placeholder="Engine oil change, filter replacement & clutch adjustment..." class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none"></textarea>
                            </div>

                            <div>
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Additional Note (Optional)</label>
                                <input type="text" id="s-note" placeholder="Call before arriving" class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none">
                            </div>

                            <!-- Service Status Switch -->
                            <div class="bg-white/80 dark:bg-slate-800/80 p-3 rounded-xl border border-amber-200 dark:border-amber-800 flex items-center justify-between">
                                <div>
                                    <p class="text-xs font-bold text-slate-800 dark:text-slate-200">Is Service Completed?</p>
                                    <p class="text-[10px] text-slate-500">Toggle if already done</p>
                                </div>
                                <label class="relative inline-flex items-center cursor-pointer">
                                    <input type="checkbox" id="s-isdone" onchange="toggleServiceDoneDateField()" class="sr-only peer">
                                    <div class="w-11 h-6 bg-slate-300 peer-focus:outline-none rounded-full peer dark:bg-slate-700 peer-checked:after:translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:left-[2px] after:bg-white after:border-slate-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-amber-600"></div>
                                </label>
                            </div>

                            <div id="service-done-date-container" class="hidden md:col-span-2">
                                <label class="block text-xs font-bold text-amber-900 dark:text-amber-300 mb-1">Service Done Date</label>
                                <input type="date" id="s-donedate" class="w-full px-3.5 py-2.5 rounded-xl border border-amber-300 dark:border-amber-700 bg-white dark:bg-slate-800 text-slate-900 dark:text-white text-sm focus:ring-2 focus:ring-amber-500 outline-none">
                            </div>
                        </div>

                        <button type="submit" class="w-full bg-amber-600 hover:bg-amber-700 text-white font-bold py-3 px-4 rounded-xl shadow-lg shadow-amber-600/30 active:scale-[0.99] transition-all flex items-center justify-center space-x-2">
                            <i class="fa-solid fa-calendar-check"></i>
                            <span>Save Service Reminder to Sheet 3</span>
                        </button>
                    </form>
                </div>
            </section>

            <!-- PANEL 4: VIEW ALL DATA & EXPORT EXCEL -->
            <section id="panel4" class="tab-content hidden space-y-4">
                <div class="bg-white dark:bg-slate-800 p-5 rounded-2xl border border-slate-200 dark:border-slate-700 shadow-sm space-y-4">
                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-100 dark:border-slate-700">
                        <div>
                            <h2 class="text-lg font-bold text-slate-900 dark:text-white">Data Viewer & Excel Export</h2>
                            <p class="text-xs text-slate-500 dark:text-slate-400">View logs and download formatted multi-sheet Excel file</p>
                        </div>

                        <!-- EXPORT BUTTON -->
                        <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold px-4 py-2.5 rounded-xl shadow-md shadow-emerald-600/20 active:scale-95 transition-all flex items-center justify-center space-x-2">
                            <i class="fa-solid fa-file-excel text-lg"></i>
                            <span>Download Excel File (.xlsx)</span>
                        </button>
                    </div>

                    <!-- TABLE SHEET SWITCHER TABS -->
                    <div class="flex border-b border-slate-200 dark:border-slate-700 space-x-2">
                        <button onclick="switchViewerTab('sheet1')" id="tab-btn-sheet1" class="viewer-tab-btn px-3 py-2 border-b-2 border-brand-500 font-bold text-xs text-brand-600 dark:text-brand-400">
                            Sheet 1: Repairs
                        </button>
                        <button onclick="switchViewerTab('sheet2')" id="tab-btn-sheet2" class="viewer-tab-btn px-3 py-2 border-b-2 border-transparent font-medium text-xs text-slate-500 hover:text-slate-800 dark:hover:text-slate-200">
                            Sheet 2: Inventory
                        </button>
                        <button onclick="switchViewerTab('sheet3')" id="tab-btn-sheet3" class="viewer-tab-btn px-3 py-2 border-b-2 border-transparent font-medium text-xs text-amber-600 dark:text-amber-400">
                            Sheet 3: Service Reminders
                        </button>
                    </div>

                    <!-- SEARCH FILTER -->
                    <div class="relative">
                        <i class="fa-solid fa-magnifying-glass absolute left-3.5 top-3 text-slate-400 text-xs"></i>
                        <input type="text" id="table-search" onkeyup="filterTables()" placeholder="Search by party name, tractor model, chassis..." class="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-300 dark:border-slate-600 bg-slate-50 dark:bg-slate-700/50 text-xs text-slate-800 dark:text-slate-100 outline-none">
                    </div>

                    <!-- SHEET 1 TABLE VIEW -->
                    <div id="view-sheet1" class="sheet-table-container overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-700">
                        <table class="w-full text-left text-xs text-slate-700 dark:text-slate-200">
                            <thead class="bg-slate-100 dark:bg-slate-700 font-semibold text-slate-900 dark:text-white">
                                <tr>
                                    <th class="p-3">Party Name</th>
                                    <th class="p-3">Chassis No</th>
                                    <th class="p-3">Model</th>
                                    <th class="p-3">Repair Date</th>
                                    <th class="p-3">Item</th>
                                    <th class="p-3">Company</th>
                                    <th class="p-3">Hours</th>
                                    <th class="p-3">Complaint</th>
                                    <th class="p-3 text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody id="tbody-sheet1" class="divide-y divide-slate-100 dark:divide-slate-700">
                                <!-- Dynamic Rows -->
                            </tbody>
                        </table>
                    </div>

                    <!-- SHEET 2 TABLE VIEW -->
                    <div id="view-sheet2" class="sheet-table-container hidden overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-700">
                        <table class="w-full text-left text-xs text-slate-700 dark:text-slate-200">
                            <thead class="bg-slate-100 dark:bg-slate-700 font-semibold text-slate-900 dark:text-white">
                                <tr>
                                    <th class="p-3">Action Type</th>
                                    <th class="p-3">Item Name</th>
                                    <th class="p-3">Company</th>
                                    <th class="p-3">Action Date</th>
                                    <th class="p-3">Chassis No</th>
                                    <th class="p-3">Model</th>
                                    <th class="p-3">Party Name</th>
                                    <th class="p-3 text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody id="tbody-sheet2" class="divide-y divide-slate-100 dark:divide-slate-700">
                                <!-- Dynamic Rows -->
                            </tbody>
                        </table>
                    </div>

                    <!-- SHEET 3 TABLE VIEW -->
                    <div id="view-sheet3" class="sheet-table-container hidden overflow-x-auto rounded-xl border border-slate-200 dark:border-slate-700">
                        <table class="w-full text-left text-xs text-slate-700 dark:text-slate-200">
                            <thead class="bg-amber-100 dark:bg-amber-900/50 font-semibold text-amber-950 dark:text-amber-200">
                                <tr>
                                    <th class="p-3">Party Name</th>
                                    <th class="p-3">Model</th>
                                    <th class="p-3">Phone</th>
                                    <th class="p-3">Due Date</th>
                                    <th class="p-3">Status</th>
                                    <th class="p-3">Done Date</th>
                                    <th class="p-3">Complaint/Work</th>
                                    <th class="p-3">Address & Note</th>
                                    <th class="p-3 text-center">Action</th>
                                </tr>
                            </thead>
                            <tbody id="tbody-sheet3" class="divide-y divide-amber-100/50 dark:divide-amber-900/30">
                                <!-- Dynamic Rows -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </section>

        </main>
    </div>

    <!-- MOBILE BOTTOM NAVIGATION (IPHONE FRIENDLY) -->
    <nav class="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-white/95 dark:bg-slate-800/95 backdrop-blur-md border-t border-slate-200 dark:border-slate-700 pb-safe px-2 py-2 flex justify-around">
        <button onclick="switchTab('panel1')" id="nav-mob-panel1" class="nav-mob-btn flex flex-col items-center p-1.5 text-brand-600 dark:text-brand-400">
            <i class="fa-solid fa-wrench text-base"></i>
            <span class="text-[10px] font-medium mt-0.5">Repair</span>
        </button>

        <button onclick="switchTab('panel2')" id="nav-mob-panel2" class="nav-mob-btn flex flex-col items-center p-1.5 text-slate-400 dark:text-slate-500">
            <i class="fa-solid fa-boxes-stacked text-base"></i>
            <span class="text-[10px] font-medium mt-0.5">Stock</span>
        </button>

        <button onclick="switchTab('panel3')" id="nav-mob-panel3" class="nav-mob-btn flex flex-col items-center p-1.5 text-slate-400 dark:text-slate-500">
            <i class="fa-solid fa-bell text-base"></i>
            <span class="text-[10px] font-medium mt-0.5">Service</span>
        </button>

        <button onclick="switchTab('panel4')" id="nav-mob-panel4" class="nav-mob-btn flex flex-col items-center p-1.5 text-slate-400 dark:text-slate-500">
            <i class="fa-solid fa-table-list text-base"></i>
            <span class="text-[10px] font-medium mt-0.5">Data/Excel</span>
        </button>
    </nav>

    <!-- TOAST NOTIFICATION -->
    <div id="toast" class="fixed bottom-16 md:bottom-6 right-4 z-50 transform translate-y-20 opacity-0 transition-all duration-300 bg-slate-900 text-white dark:bg-white dark:text-slate-900 px-4 py-3 rounded-xl shadow-xl flex items-center space-x-2 text-xs font-semibold">
        <i id="toast-icon" class="fa-solid fa-circle-check text-emerald-400 dark:text-emerald-600 text-base"></i>
        <span id="toast-msg">Action performed successfully!</span>
    </div>

    <!-- MAIN APP SCRIPT & CLOUD SYNC LOGIC -->
    <script>
        // Company options logic
        const COMPANIES = {
            'Battery': ['Amaron', 'Exide', 'Livguard'],
            'Self Starter': ['Autoluck', 'Spark Minda'],
            'Alternator': ['Autoluck', 'Spark Minda']
        };

        // Application State
        let dbData = {
            sheet1: [], // Repair logs
            sheet2: [], // Inventory logs
            sheet3: []  // Service logs
        };

        // Initialize App
        window.addEventListener('DOMContentLoaded', () => {
            // Set today's default dates
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('r-date').value = today;
            document.getElementById('i-date').value = today;
            document.getElementById('s-duedate').value = today;
            document.getElementById('s-donedate').value = today;

            // Load company options
            updateCompanyOptions('r-category', 'r-company');
            updateCompanyOptions('i-category', 'i-company');

            // Load local storage initial state
            loadFromLocalStorage();

            // Dark Mode Setup
            if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
                document.documentElement.classList.add('dark');
            }
        });

        // Theme Toggle Function
        document.getElementById('theme-toggle').addEventListener('click', () => {
            document.documentElement.classList.toggle('dark');
        });

        // Dynamic Company Dropdowns
        function updateCompanyOptions(catSelectId, compSelectId) {
            const category = document.getElementById(catSelectId).value;
            const compSelect = document.getElementById(compSelectId);
            compSelect.innerHTML = '';
            
            const options = COMPANIES[category] || [];
            options.forEach(company => {
                const opt = document.createElement('option');
                opt.value = company;
                opt.textContent = company;
                compSelect.appendChild(opt);
            });
        }

        // Toggle Inventory Fitting Conditional Fields
        function toggleInventoryFields() {
            const invType = document.querySelector('input[name="inv-type"]:checked').value;
            const fittingFields = document.getElementById('inv-fitting-fields');
            if (invType === 'Fitted/Out') {
                fittingFields.classList.remove('hidden');
            } else {
                fittingFields.classList.add('hidden');
            }
        }

        // Toggle Service Done Date
        function toggleServiceDoneDateField() {
            const isDone = document.getElementById('s-isdone').checked;
            const doneContainer = document.getElementById('service-done-date-container');
            if (isDone) {
                doneContainer.classList.remove('hidden');
            } else {
                doneContainer.classList.add('hidden');
            }
        }

        // Tab Switcher
        function switchTab(panelId) {
            // Hide all tab content
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.getElementById(panelId).classList.remove('hidden');

            // Update Desktop Nav
            document.querySelectorAll('.nav-desk-btn').forEach(btn => {
                btn.className = 'nav-desk-btn flex items-center space-x-3 w-full px-3 py-2.5 rounded-xl font-medium text-sm text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-700 transition-all';
            });
            const activeDeskBtn = document.getElementById(`nav-desk-${panelId}`);
            if (activeDeskBtn) {
                if (panelId === 'panel3') {
                    activeDeskBtn.className = 'nav-desk-btn flex items-center space-x-3 w-full px-3 py-2.5 rounded-xl font-bold text-sm text-amber-600 bg-amber-50 dark:bg-amber-900/30 dark:text-amber-400 transition-all';
                } else {
                    activeDeskBtn.className = 'nav-desk-btn flex items-center space-x-3 w-full px-3 py-2.5 rounded-xl font-bold text-sm text-brand-600 bg-brand-50 dark:bg-brand-900/30 dark:text-brand-400 transition-all';
                }
            }

            // Update Mobile Nav
            document.querySelectorAll('.nav-mob-btn').forEach(btn => {
                btn.className = 'nav-mob-btn flex flex-col items-center p-1.5 text-slate-400 dark:text-slate-500';
            });
            const activeMobBtn = document.getElementById(`nav-mob-${panelId}`);
            if (activeMobBtn) {
                activeMobBtn.className = panelId === 'panel3' 
                    ? 'nav-mob-btn flex flex-col items-center p-1.5 text-amber-600 dark:text-amber-400' 
                    : 'nav-mob-btn flex flex-col items-center p-1.5 text-brand-600 dark:text-brand-400';
            }

            // Scroll to top
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        // Form Submit: Repair (Sheet 1)
        function handleRepairSubmit(e) {
            e.preventDefault();
            const record = {
                id: Date.now().toString(),
                partyName: document.getElementById('r-party').value,
                chassisNo: document.getElementById('r-chassis').value,
                tractorModel: document.getElementById('r-model').value,
                repairDate: document.getElementById('r-date').value,
                itemCategory: document.getElementById('r-category').value,
                companyName: document.getElementById('r-company').value,
                tractorHours: document.getElementById('r-hours').value || '0',
                complaint: document.getElementById('r-complaint').value
            };

            dbData.sheet1.unshift(record);
            saveAndSync();
            document.getElementById('form-repair').reset();
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('r-date').value = today;
            updateCompanyOptions('r-category', 'r-company');
            showToast('Repair entry saved to Sheet 1!');
        }

        // Form Submit: Inventory (Sheet 2)
        function handleInventorySubmit(e) {
            e.preventDefault();
            const invType = document.querySelector('input[name="inv-type"]:checked').value;
            const record = {
                id: Date.now().toString(),
                entryType: invType,
                itemCategory: document.getElementById('i-category').value,
                companyName: document.getElementById('i-company').value,
                actionDate: document.getElementById('i-date').value,
                chassisNo: invType === 'Fitted/Out' ? (document.getElementById('i-chassis').value || 'N/A') : 'N/A (Stock In)',
                tractorModel: invType === 'Fitted/Out' ? (document.getElementById('i-model').value || 'N/A') : 'N/A (Stock In)',
                partyName: invType === 'Fitted/Out' ? (document.getElementById('i-party').value || 'N/A') : 'Store Stock'
            };

            dbData.sheet2.unshift(record);
            saveAndSync();
            document.getElementById('form-inventory').reset();
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('i-date').value = today;
            updateCompanyOptions('i-category', 'i-company');
            toggleInventoryFields();
            showToast('Inventory record saved to Sheet 2!');
        }

        // Form Submit: Service Reminder (Sheet 3)
        function handleServiceSubmit(e) {
            e.preventDefault();
            const isDone = document.getElementById('s-isdone').checked;
            const record = {
                id: Date.now().toString(),
                partyName: document.getElementById('s-party').value,
                tractorModel: document.getElementById('s-model').value,
                phone: document.getElementById('s-phone').value,
                address: document.getElementById('s-address').value || 'N/A',
                note: document.getElementById('s-note').value || 'N/A',
                serviceDueDate: document.getElementById('s-duedate').value,
                complaint: document.getElementById('s-complaint').value,
                status: isDone ? 'Done' : 'Pending',
                serviceDoneDate: isDone ? document.getElementById('s-donedate').value : 'Pending'
            };

            dbData.sheet3.unshift(record);
            saveAndSync();
            document.getElementById('form-service').reset();
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('s-duedate').value = today;
            document.getElementById('s-donedate').value = today;
            toggleServiceDoneDateField();
            showToast('Tractor service reminder saved to Sheet 3!');
        }

        // Delete Row Item
        function deleteRecord(sheetKey, id) {
            if(confirm('Are you sure you want to delete this record?')) {
                dbData[sheetKey] = dbData[sheetKey].filter(item => item.id !== id);
                saveAndSync();
                showToast('Record deleted!');
            }
        }

        // Toggle Service Status from Table
        function toggleServiceStatus(id) {
            const item = dbData.sheet3.find(i => i.id === id);
            if (item) {
                if (item.status === 'Pending') {
                    item.status = 'Done';
                    item.serviceDoneDate = new Date().toISOString().split('T')[0];
                } else {
                    item.status = 'Pending';
                    item.serviceDoneDate = 'Pending';
                }
                saveAndSync();
                showToast('Service status updated!');
            }
        }

        // Render Tables & Update Badges
        function renderTables() {
            // Update Quick Stats
            document.getElementById('stat-repairs').textContent = dbData.sheet1.length;
            document.getElementById('stat-inventory').textContent = dbData.sheet2.length;
            document.getElementById('stat-services').textContent = dbData.sheet3.filter(s => s.status === 'Pending').length;

            // Render Sheet 1
            const tbody1 = document.getElementById('tbody-sheet1');
            tbody1.innerHTML = dbData.sheet1.length ? dbData.sheet1.map(r => `
                <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td class="p-3 font-semibold text-slate-900 dark:text-white">${escapeHtml(r.partyName)}</td>
                    <td class="p-3">${escapeHtml(r.chassisNo)}</td>
                    <td class="p-3">${escapeHtml(r.tractorModel)}</td>
                    <td class="p-3">${escapeHtml(r.repairDate)}</td>
                    <td class="p-3"><span class="px-2 py-0.5 rounded-md bg-blue-100 text-blue-800 dark:bg-blue-900/40 dark:text-blue-300 text-[10px] font-bold">${escapeHtml(r.itemCategory)}</span></td>
                    <td class="p-3 font-medium">${escapeHtml(r.companyName)}</td>
                    <td class="p-3">${escapeHtml(r.tractorHours)} hrs</td>
                    <td class="p-3 max-w-xs truncate">${escapeHtml(r.complaint)}</td>
                    <td class="p-3 text-center">
                        <button onclick="deleteRecord('sheet1', '${r.id}')" class="text-rose-500 hover:text-rose-700 p-1"><i class="fa-solid fa-trash"></i></button>
                    </td>
                </tr>
            `).join('') : '<tr><td colspan="9" class="p-4 text-center text-slate-400">No repair logs added yet.</td></tr>';

            // Render Sheet 2
            const tbody2 = document.getElementById('tbody-sheet2');
            tbody2.innerHTML = dbData.sheet2.length ? dbData.sheet2.map(i => `
                <tr class="hover:bg-slate-50 dark:hover:bg-slate-700/30">
                    <td class="p-3">
                        <span class="px-2 py-0.5 rounded-md text-[10px] font-bold ${i.entryType === 'Stock In' ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/40 dark:text-emerald-300' : 'bg-orange-100 text-orange-800 dark:bg-orange-900/40 dark:text-orange-300'}">
                            ${escapeHtml(i.entryType)}
                        </span>
                    </td>
                    <td class="p-3 font-semibold">${escapeHtml(i.itemCategory)}</td>
                    <td class="p-3">${escapeHtml(i.companyName)}</td>
                    <td class="p-3">${escapeHtml(i.actionDate)}</td>
                    <td class="p-3">${escapeHtml(i.chassisNo)}</td>
                    <td class="p-3">${escapeHtml(i.tractorModel)}</td>
                    <td class="p-3">${escapeHtml(i.partyName)}</td>
                    <td class="p-3 text-center">
                        <button onclick="deleteRecord('sheet2', '${i.id}')" class="text-rose-500 hover:text-rose-700 p-1"><i class="fa-solid fa-trash"></i></button>
                    </td>
                </tr>
            `).join('') : '<tr><td colspan="8" class="p-4 text-center text-slate-400">No inventory logs added yet.</td></tr>';

            // Render Sheet 3
            const tbody3 = document.getElementById('tbody-sheet3');
            tbody3.innerHTML = dbData.sheet3.length ? dbData.sheet3.map(s => `
                <tr class="hover:bg-amber-50/50 dark:hover:bg-amber-900/20">
                    <td class="p-3 font-semibold text-slate-900 dark:text-white">${escapeHtml(s.partyName)}</td>
                    <td class="p-3">${escapeHtml(s.tractorModel)}</td>
                    <td class="p-3"><a href="tel:${s.phone}" class="text-brand-600 dark:text-brand-400 underline font-medium"><i class="fa-solid fa-phone text-[10px]"></i> ${escapeHtml(s.phone)}</a></td>
                    <td class="p-3 font-bold text-amber-700 dark:text-amber-400">${escapeHtml(s.serviceDueDate)}</td>
                    <td class="p-3">
                        <button onclick="toggleServiceStatus('${s.id}')" class="px-2 py-1 rounded-lg text-[10px] font-bold cursor-pointer transition-all ${s.status === 'Done' ? 'bg-emerald-100 text-emerald-800 dark:bg-emerald-900/50 dark:text-emerald-300' : 'bg-rose-100 text-rose-800 dark:bg-rose-900/50 dark:text-rose-300'}">
                            <i class="fa-solid ${s.status === 'Done' ? 'fa-check-circle' : 'fa-hourglass-half'} mr-1"></i>${escapeHtml(s.status)}
                        </button>
                    </td>
                    <td class="p-3">${escapeHtml(s.serviceDoneDate)}</td>
                    <td class="p-3 max-w-xs truncate">${escapeHtml(s.complaint)}</td>
                    <td class="p-3 text-[11px] text-slate-500">${escapeHtml(s.address)} <br> <span class="italic text-slate-400">${escapeHtml(s.note)}</span></td>
                    <td class="p-3 text-center">
                        <button onclick="deleteRecord('sheet3', '${s.id}')" class="text-rose-500 hover:text-rose-700 p-1"><i class="fa-solid fa-trash"></i></button>
                    </td>
                </tr>
            `).join('') : '<tr><td colspan="9" class="p-4 text-center text-slate-400">No service reminder logs added yet.</td></tr>';
        }

        // Viewer Sheet Tab Switcher
        function switchViewerTab(sheetId) {
            document.querySelectorAll('.sheet-table-container').forEach(c => c.classList.add('hidden'));
            document.getElementById(`view-${sheetId}`).classList.remove('hidden');

            document.querySelectorAll('.viewer-tab-btn').forEach(btn => {
                btn.className = 'viewer-tab-btn px-3 py-2 border-b-2 border-transparent font-medium text-xs text-slate-500 hover:text-slate-800 dark:hover:text-slate-200';
            });
            document.getElementById(`tab-btn-${sheetId}`).className = 'viewer-tab-btn px-3 py-2 border-b-2 border-brand-500 font-bold text-xs text-brand-600 dark:text-brand-400';
        }

        // Filter Tables Search
        function filterTables() {
            const query = document.getElementById('table-search').value.toLowerCase();
            document.querySelectorAll('tbody tr').forEach(row => {
                const text = row.textContent.toLowerCase();
                row.style.display = text.includes(query) ? '' : 'none';
            });
        }

        // Save Data & Trigger Sync
        function saveAndSync() {
            // Save locally
            localStorage.setItem('spare_app_data', JSON.stringify(dbData));
            renderTables();

            // Broadcast channel for instantaneous cross-tab / mobile updates
            if (window.BroadcastChannel) {
                const bc = new BroadcastChannel('spare_app_channel');
                bc.postMessage(dbData);
            }
        }

        // Load local storage
        function loadFromLocalStorage() {
            const local = localStorage.getItem('spare_app_data');
            if (local) {
                try {
                    dbData = JSON.parse(local);
                } catch(e) {
                    console.error("Storage error", e);
                }
            }
            renderTables();
        }

        // Listen for live broadcast events across tabs/browser instances
        if (window.BroadcastChannel) {
            const bc = new BroadcastChannel('spare_app_channel');
            bc.onmessage = (event) => {
                dbData = event.data;
                renderTables();
            };
        }

        // EXCEL MULTI-SHEET EXPORT FUNCTION
        function exportToExcel() {
            const wb = XLSX.utils.book_new();

            // Format Sheet 1
            const sheet1Data = dbData.sheet1.map(r => ({
                "Party Name": r.partyName,
                "Chassis Number": r.chassisNo,
                "Tractor Model": r.tractorModel,
                "Repair Date": r.repairDate,
                "Complaint": r.complaint,
                "Tractor Hours": r.tractorHours,
                "Item Category": r.itemCategory,
                "Company Name": r.companyName
            }));
            const ws1 = XLSX.utils.json_to_sheet(sheet1Data);
            XLSX.utils.book_append_sheet(wb, ws1, "Service_Repair_Logs");

            // Format Sheet 2
            const sheet2Data = dbData.sheet2.map(i => ({
                "Entry Type": i.entryType,
                "Item Name": i.itemCategory,
                "Company Name": i.companyName,
                "Action Date": i.actionDate,
                "Chassis Number": i.chassisNo,
                "Tractor Model": i.tractorModel,
                "Party Name": i.partyName
            }));
            const ws2 = XLSX.utils.json_to_sheet(sheet2Data);
            XLSX.utils.book_append_sheet(wb, ws2, "Inventory_Logs");

            // Format Sheet 3
            const sheet3Data = dbData.sheet3.map(s => ({
                "Party Name": s.partyName,
                "Tractor Model": s.tractorModel,
                "Phone Number": s.phone,
                "Address": s.address,
                "Note": s.note,
                "Service Due Date": s.serviceDueDate,
                "Complaint / Work Details": s.complaint,
                "Status": s.status,
                "Service Done Date": s.serviceDoneDate
            }));
            const ws3 = XLSX.utils.json_to_sheet(sheet3Data);
            XLSX.utils.book_append_sheet(wb, ws3, "Tractor_Service_Logs");

            // Save File
            XLSX.writeFile(wb, "Spare_And_Tractor_Service_Data.xlsx");
            showToast("Excel file downloaded!");
        }

        // Toast Helper
        function showToast(msg) {
            const toast = document.getElementById('toast');
            document.getElementById('toast-msg').textContent = msg;
            toast.classList.remove('translate-y-20', 'opacity-0');
            setTimeout(() => {
                toast.classList.add('translate-y-20', 'opacity-0');
            }, 3000);
        }

        // Utility: Escape HTML
        function escapeHtml(str) {
            if (!str) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }
    </script>
</body>
</html>
