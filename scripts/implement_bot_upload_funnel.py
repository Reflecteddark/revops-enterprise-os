# -*- coding: utf-8 -*-
"""
Implements:
1. Audio drag-and-drop attachment field in #leadForm on website
2. Step 2 Screen in #formSuccess with direct 3-calls upload buttons to MAX Bot and Telegram Bot
3. Client-side uploader JS logic with file preview and remove buttons
"""

import os
import re

repo = r"C:\Users\strel\.gemini\antigravity\scratch\revops-enterprise-os"
index_path = os.path.join(repo, "index.html")

with open(index_path, "r", encoding="utf-8") as f:
    content = f.read()

# ----------------------------------------------------
# 1. AUDIO DRAG & DROP UPLOADER FIELD IN #leadForm
# ----------------------------------------------------
uploader_html = """
                        <!-- AUDIO FILES ATTACHMENT FIELD (OPTIONAL 3 CALLS UPLOAD) -->
                        <div>
                            <div class="flex items-center justify-between mb-2">
                                <label class="block text-xs font-semibold text-slate-300 uppercase tracking-wider">Прикрепите 3 звонка прямо здесь (по желанию):</label>
                                <span class="text-[10px] font-mono text-cyan-400">до 3 файлов (mp3, wav, m4a)</span>
                            </div>
                            <div id="audioDropArea" class="border-2 border-dashed border-slate-700 hover:border-emerald-500/60 rounded-2xl p-4 text-center cursor-pointer transition bg-slate-950/60 group">
                                <input type="file" id="audioFileInput" multiple accept=".mp3,.wav,.m4a,.ogg,.aac,audio/*" class="hidden">
                                <div class="flex flex-col items-center justify-center gap-1.5 pointer-events-none" id="dropAreaPrompt">
                                    <div class="w-10 h-10 rounded-xl bg-slate-900 border border-slate-800 flex items-center justify-center text-emerald-400 text-lg group-hover:scale-110 transition-transform">
                                        <i class="fas fa-cloud-arrow-up"></i>
                                    </div>
                                    <div class="text-xs text-slate-300 font-medium">
                                        <span class="text-emerald-400 font-bold underline">Выберите до 3 файлов</span> или перетащите их сюда
                                    </div>
                                    <div class="text-[10px] text-slate-500 font-mono">
                                        Если записей нет под рукой — отправьте их в MAX или Telegram на следующем шаге
                                    </div>
                                </div>
                                <div id="selectedAudioList" class="hidden mt-3 space-y-1.5 text-left border-t border-slate-800/80 pt-2 text-xs"></div>
                            </div>
                        </div>
"""

# Insert uploader right before the consent checkbox
if 'id="audioDropArea"' not in content:
    content = content.replace(
        '<div class="flex items-start gap-2.5 pt-1 text-left">',
        uploader_html + '\n                        <div class="flex items-start gap-2.5 pt-1 text-left">'
    )
    print("✓ Audio drag & drop uploader added to #leadForm")

# ----------------------------------------------------
# 2. UPGRADE #formSuccess INTO STEP 2 BOT UPLOAD ACTION HUB
# ----------------------------------------------------
old_success_pattern = r'<div id="formSuccess" class="hidden text-center py-6" style="display: none;">.*?</div>\s*</div>\s*</div>\s*</div>\s*</section>'

new_success_html = """<div id="formSuccess" class="hidden text-center py-4" style="display: none;">
                        <!-- Step Badge -->
                        <div class="inline-flex items-center gap-2 px-3.5 py-1 rounded-full bg-emerald-500/10 text-emerald-400 font-mono text-xs font-bold border border-emerald-500/20 mb-3">
                            <span class="w-2 h-2 rounded-full bg-emerald-400 animate-pulse"></span>
                            ШАГ 2 ИЗ 2 • ОТПРАВКА 3 ЗВОНКОВ
                        </div>
                        
                        <h3 class="text-2xl sm:text-3xl font-extrabold text-white mb-2 font-display">
                            Отправьте 3 аудиозаписи в бот для старта ИИ-анализа
                        </h3>
                        <p id="formSuccessSubtitle" class="text-slate-300 text-xs sm:text-sm max-w-xl mx-auto mb-6">
                            Ваша заявка зафиксирована! Чтобы нейросеть Whisper Pro оцифровала сливы выручки — прикрепите 3 вчерашних звонка в удобный мессенджер (файлами или перешлите аудио):
                        </p>

                        <!-- Two Main Action Bot Cards (MAX + Telegram) -->
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3.5 mb-6 text-left">
                            <!-- MAX Bot Card -->
                            <a href="https://max.ru/se14526668_bot" target="_blank" rel="noopener noreferrer" class="p-5 rounded-2xl bg-purple-950/40 hover:bg-purple-900/60 border-2 border-purple-500/50 hover:border-purple-400 transition-all shadow-xl shadow-purple-500/15 group flex flex-col justify-between">
                                <div>
                                    <div class="flex items-center justify-between mb-3">
                                        <span class="px-2.5 py-0.5 rounded-md bg-purple-500/20 text-purple-300 font-mono text-[10px] font-bold">100% БЕЗ VPN В РФ</span>
                                        <i class="fas fa-arrow-right text-purple-400 group-hover:translate-x-1 transition-transform"></i>
                                    </div>
                                    <div class="font-display font-bold text-white text-base sm:text-lg mb-1 flex items-center gap-2">
                                        <i class="fas fa-comment-dots text-purple-400"></i>
                                        <span>Отправить в MAX-бот</span>
                                    </div>
                                    <p class="text-xs text-slate-300 leading-relaxed">
                                        Откройте диалог с ботом @se14526668_bot и прикрепите 3 файла звонков. Работает молниеносно на любом операторе РФ.
                                    </p>
                                </div>
                                <div class="mt-4 pt-3 border-t border-purple-500/30 text-xs font-bold text-purple-300 flex items-center gap-1.5">
                                    <span>Загрузить 3 звонка через MAX</span> →
                                </div>
                            </a>

                            <!-- Telegram Bot Card -->
                            <a href="https://t.me/RevOps_Super_Audit_Bot?start=audit3" target="_blank" rel="noopener noreferrer" class="p-5 rounded-2xl bg-sky-950/40 hover:bg-sky-900/60 border-2 border-sky-500/50 hover:border-sky-400 transition-all shadow-xl shadow-sky-500/15 group flex flex-col justify-between">
                                <div>
                                    <div class="flex items-center justify-between mb-3">
                                        <span class="px-2.5 py-0.5 rounded-md bg-sky-500/20 text-sky-300 font-mono text-[10px] font-bold">TELEGRAM СЕТЬ</span>
                                        <i class="fas fa-arrow-right text-sky-400 group-hover:translate-x-1 transition-transform"></i>
                                    </div>
                                    <div class="font-display font-bold text-white text-base sm:text-lg mb-1 flex items-center gap-2">
                                        <i class="fab fa-telegram-plane text-sky-400"></i>
                                        <span>Отправить в Telegram-бот</span>
                                    </div>
                                    <p class="text-xs text-slate-300 leading-relaxed">
                                        Перешлите 3 звонка боту @RevOps_Super_Audit_Bot. Бот подтвердит получение каждого файла и передаст на разбор.
                                    </p>
                                </div>
                                <div class="mt-4 pt-3 border-t border-sky-500/30 text-xs font-bold text-sky-300 flex items-center gap-1.5">
                                    <span>Загрузить 3 звонка через TG</span> →
                                </div>
                            </a>
                        </div>

                        <!-- Direct Founder Alternative -->
                        <div class="p-4 rounded-2xl bg-slate-950/80 border border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-3 text-xs text-slate-300">
                            <span class="flex items-center gap-2.5">
                                <img src="images/founder-dmitry.jpg" alt="Дмитрий" class="w-7 h-7 rounded-full object-cover border border-emerald-400">
                                <span>Или отправьте 3 аудиозаписи лично основателю Дмитрию Федотову:</span>
                            </span>
                            <div class="flex items-center gap-3 font-semibold">
                                <a href="https://t.me/dm1918" target="_blank" class="text-cyan-400 hover:underline flex items-center gap-1">
                                    <i class="fab fa-telegram"></i> @dm1918 (Telegram)
                                </a>
                                <span>•</span>
                                <a href="https://max.ru/u/f9LHodD0cOLRcNKRakz94FpZjmUJ25lYzJUMzPBJyLKwxM1Tzzai1aF-dTg" target="_blank" class="text-purple-400 hover:underline flex items-center gap-1">
                                    <i class="fas fa-comment-dots"></i> Профиль в MAX
                                </a>
                            </div>
                        </div>
                        <p class="text-xs text-slate-500 mt-4">⏱️ Время оцифровки звонков и выдачи PDF-карты: 20 минут. Без спама и звонков.</p>
                    </div>
                </div>
            </div>
        </div>
    </section>"""

if re.search(old_success_pattern, content, flags=re.DOTALL):
    content = re.sub(old_success_pattern, new_success_html, content, flags=re.DOTALL, count=1)
    print("✓ #formSuccess transformed into Step 2 Bot Upload Action Hub")
else:
    print("⚠ Notice: old_success_pattern not found directly, performing targeted replacement")
    if '<div id="formSuccess"' in content:
        start_idx = content.find('<div id="formSuccess"')
        end_idx = content.find('</section>', start_idx)
        if start_idx != -1 and end_idx != -1:
            content = content[:start_idx] + new_success_html + content[end_idx + len('</section>'):]
            print("✓ Targeted index-based replacement for #formSuccess succeeded")

# ----------------------------------------------------
# 3. CLIENT-SIDE JS LOGIC FOR AUDIO UPLOADER
# ----------------------------------------------------
uploader_js = """
        // ========================================================
        // 5. AUDIO DRAG & DROP UPLOADER LOGIC
        // ========================================================
        (function() {
            const dropArea = document.getElementById('audioDropArea');
            const fileInput = document.getElementById('audioFileInput');
            const listContainer = document.getElementById('selectedAudioList');
            let selectedFiles = [];

            if (!dropArea || !fileInput || !listContainer) return;

            dropArea.addEventListener('click', (e) => {
                if (e.target.closest('.remove-file-btn')) return;
                fileInput.click();
            });

            ['dragenter', 'dragover'].forEach(name => {
                dropArea.addEventListener(name, (e) => {
                    e.preventDefault();
                    dropArea.classList.add('border-emerald-400', 'bg-slate-900');
                });
            });

            ['dragleave', 'drop'].forEach(name => {
                dropArea.addEventListener(name, (e) => {
                    e.preventDefault();
                    dropArea.classList.remove('border-emerald-400', 'bg-slate-900');
                });
            });

            dropArea.addEventListener('drop', (e) => {
                const dt = e.dataTransfer;
                if (dt && dt.files && dt.files.length > 0) {
                    handleFiles(dt.files);
                }
            });

            fileInput.addEventListener('change', () => {
                if (fileInput.files && fileInput.files.length > 0) {
                    handleFiles(fileInput.files);
                }
            });

            function handleFiles(files) {
                for (let i = 0; i < files.length; i++) {
                    if (selectedFiles.length >= 3) break;
                    selectedFiles.push(files[i]);
                }
                renderFileList();
            }

            function renderFileList() {
                if (selectedFiles.length === 0) {
                    listContainer.classList.add('hidden');
                    listContainer.innerHTML = '';
                    return;
                }
                listContainer.classList.remove('hidden');
                listContainer.innerHTML = '';
                
                selectedFiles.forEach((file, index) => {
                    const sizeMb = (file.size / (1024 * 1024)).toFixed(2);
                    const item = document.createElement('div');
                    item.className = 'flex items-center justify-between p-2 rounded-xl bg-slate-900 border border-slate-800 text-slate-200';
                    item.innerHTML = `
                        <div class="flex items-center gap-2 truncate pr-2">
                            <i class="fas fa-file-audio text-emerald-400 text-sm shrink-0"></i>
                            <span class="truncate font-mono text-[11px]">${file.name}</span>
                            <span class="text-[10px] text-slate-500 font-mono">(${sizeMb} MB)</span>
                        </div>
                        <button type="button" data-idx="${index}" class="remove-file-btn text-slate-500 hover:text-rose-400 p-1 text-xs transition">
                            <i class="fas fa-times"></i>
                        </button>
                    `;
                    listContainer.appendChild(item);
                });

                // Attach remove handlers
                const removeBtns = listContainer.querySelectorAll('.remove-file-btn');
                removeBtns.forEach(btn => {
                    btn.addEventListener('click', (e) => {
                        e.stopPropagation();
                        const idx = parseInt(btn.getAttribute('data-idx'));
                        selectedFiles.splice(idx, 1);
                        renderFileList();
                    });
                });
            }

            window.getSelectedAudioCount = function() {
                return selectedFiles.length;
            };
        })();
"""

content = content.replace('</body>', '<script>' + uploader_js + '</script>\n</body>')

with open(index_path, "w", encoding="utf-8") as f:
    f.write(content)

print(f"✓ index.html successfully upgraded! ({len(content)} bytes)")
