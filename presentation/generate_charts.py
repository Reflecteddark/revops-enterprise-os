import os
import sys

sys.stdout.reconfigure(encoding='utf-8')
os.makedirs("presentation/charts", exist_ok=True)

# 1. FUNNEL CHART SVG (Horizontal funnel with SLA breach highlight)
funnel_svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 540 260" width="100%" height="100%">
  <defs>
    <linearGradient id="gradRed" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0%" stop-color="#DC2626"/>
      <stop offset="100%" stop-color="#EF4444"/>
    </linearGradient>
    <linearGradient id="gradGreen" x1="0" y1="0" x2="1" y2="0">
      <stop offset="0%" stop-color="#15803D"/>
      <stop offset="100%" stop-color="#22C55E"/>
    </linearGradient>
  </defs>

  <!-- Stage 1 -->
  <g transform="translate(10, 15)">
    <text x="0" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#1E293B">1. Новый лид</text>
    <rect x="150" y="2" width="220" height="20" rx="4" fill="url(#gradGreen)"/>
    <text x="380" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A">950 000 ₽</text>
    <text x="475" y="16" font-family="Inter, sans-serif" font-size="10" fill="#15803D" font-weight="600">0ч / 24ч (OK)</text>
  </g>

  <!-- Stage 2 -->
  <g transform="translate(10, 48)">
    <text x="0" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#1E293B">2. Квалификация</text>
    <rect x="150" y="2" width="160" height="20" rx="4" fill="url(#gradGreen)"/>
    <text x="320" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A">120 000 ₽</text>
    <text x="475" y="16" font-family="Inter, sans-serif" font-size="10" fill="#15803D" font-weight="600">24ч / 48ч (OK)</text>
  </g>

  <!-- Stage 3 -->
  <g transform="translate(10, 81)">
    <text x="0" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#1E293B">3. Демо / Встреча</text>
    <rect x="150" y="2" width="190" height="20" rx="4" fill="url(#gradGreen)"/>
    <text x="350" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A">280 000 ₽</text>
    <text x="475" y="16" font-family="Inter, sans-serif" font-size="10" fill="#15803D" font-weight="600">72ч / 120ч (OK)</text>
  </g>

  <!-- Stage 4: SLA BREACH (Highlight) -->
  <g transform="translate(10, 114)">
    <rect x="-6" y="-2" width="532" height="28" rx="6" fill="#FEF2F2" stroke="#FCA5A5" stroke-width="1.5"/>
    <text x="0" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#991B1B">4. Согласование КП</text>
    <rect x="150" y="2" width="280" height="20" rx="4" fill="url(#gradRed)"/>
    <text x="435" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="800" fill="#DC2626">3 850 000 ₽</text>
    <text x="160" y="16" font-family="Inter, sans-serif" font-size="10" font-weight="800" fill="#FFFFFF">312 ч (НОРМА 48 ч — СРЫВ 6.5x!)</text>
  </g>

  <!-- Stage 5: SLA BREACH (Highlight) -->
  <g transform="translate(10, 147)">
    <rect x="-6" y="-2" width="532" height="28" rx="6" fill="#FEF2F2" stroke="#FCA5A5" stroke-width="1.5"/>
    <text x="0" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#991B1B">5. Счет на оплате</text>
    <rect x="150" y="2" width="180" height="20" rx="4" fill="url(#gradRed)"/>
    <text x="340" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="800" fill="#DC2626">330 000 ₽</text>
    <text x="160" y="16" font-family="Inter, sans-serif" font-size="10" font-weight="800" fill="#FFFFFF">708 ч (НОРМА 72 ч — СРЫВ 9.8x!)</text>
  </g>

  <!-- Stage 6: WON -->
  <g transform="translate(10, 180)">
    <text x="0" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#1E293B">6. Оплачено (Won)</text>
    <rect x="150" y="2" width="240" height="20" rx="4" fill="#047857"/>
    <text x="400" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#047857">1 650 000 ₽</text>
    <text x="475" y="16" font-family="Inter, sans-serif" font-size="10" fill="#047857" font-weight="600">Касса (Факт)</text>
  </g>

  <!-- Stage 7: LOST -->
  <g transform="translate(10, 213)">
    <text x="0" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#64748B">7. Архив / Отказ</text>
    <rect x="150" y="2" width="150" height="20" rx="4" fill="#94A3B8"/>
    <text x="310" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#64748B">1 000 000 ₽</text>
    <text x="475" y="16" font-family="Inter, sans-serif" font-size="10" fill="#64748B" font-weight="600">AI Recovery</text>
  </g>
</svg>"""

with open("presentation/charts/funnel.svg", "w", encoding="utf-8") as f:
    f.write(funnel_svg)

# 2. RISKS BAR CHART SVG
risks_svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 520 220" width="100%" height="100%">
  <!-- Grid lines -->
  <line x1="120" y1="20" x2="120" y2="180" stroke="#CBD5E1" stroke-width="1"/>
  <line x1="210" y1="20" x2="210" y2="180" stroke="#F1F5F9" stroke-width="1"/>
  <line x1="300" y1="20" x2="300" y2="180" stroke="#F1F5F9" stroke-width="1"/>
  <line x1="390" y1="20" x2="390" y2="180" stroke="#F1F5F9" stroke-width="1"/>
  <line x1="480" y1="20" x2="480" y2="180" stroke="#F1F5F9" stroke-width="1"/>

  <!-- D-107 -->
  <g transform="translate(10, 25)">
    <text x="105" y="15" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A" text-anchor="end">D-107 (ТехноПром)</text>
    <rect x="120" y="2" width="360" height="18" rx="3" fill="#DC2626"/>
    <text x="485" y="15" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#DC2626">3 000 000 ₽</text>
  </g>

  <!-- D-104 -->
  <g transform="translate(10, 58)">
    <text x="105" y="15" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A" text-anchor="end">D-104 (Вектор Плюс)</text>
    <rect x="120" y="2" width="130" height="18" rx="3" fill="#EF4444"/>
    <text x="258" y="15" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#DC2626">850 000 ₽</text>
  </g>

  <!-- D-106 -->
  <g transform="translate(10, 91)">
    <text x="105" y="15" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#0F172A" text-anchor="end">D-106 (СнабСервис)</text>
    <rect x="120" y="2" width="38" height="18" rx="3" fill="#F59E0B"/>
    <text x="165" y="15" font-family="Inter, sans-serif" font-size="10.5" font-weight="700" fill="#B45309">180 000 ₽</text>
  </g>

  <!-- D-103 -->
  <g transform="translate(10, 124)">
    <text x="105" y="15" font-family="Inter, sans-serif" font-size="11" font-weight="600" fill="#0F172A" text-anchor="end">D-103 (Смирнов)</text>
    <rect x="120" y="2" width="32" height="18" rx="3" fill="#F59E0B"/>
    <text x="160" y="15" font-family="Inter, sans-serif" font-size="10.5" font-weight="700" fill="#B45309">150 000 ₽</text>
  </g>

  <!-- D-109 -->
  <g transform="translate(10, 157)">
    <text x="105" y="15" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A" text-anchor="end">D-109 (ПромКомплект)</text>
    <rect x="120" y="2" width="28" height="18" rx="3" fill="#DC2626"/>
    <text x="156" y="15" font-family="Inter, sans-serif" font-size="10.5" font-weight="700" fill="#DC2626">120 000 ₽</text>
  </g>

  <!-- Axis labels -->
  <text x="120" y="198" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="middle">0</text>
  <text x="240" y="198" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="middle">1.0M ₽</text>
  <text x="360" y="198" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="middle">2.0M ₽</text>
  <text x="480" y="198" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="middle">3.0M ₽</text>
</svg>"""

with open("presentation/charts/risks_bar.svg", "w", encoding="utf-8") as f:
    f.write(risks_svg)

# 3. SPEECH AUDIT BAR CHART SVG
speech_svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 460 210" width="100%" height="100%">
  <!-- Y-Axis Lines -->
  <line x1="50" y1="20" x2="430" y2="20" stroke="#F1F5F9" stroke-width="1"/>
  <text x="40" y="24" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="end">13</text>
  
  <line x1="50" y1="65" x2="430" y2="65" stroke="#F1F5F9" stroke-width="1"/>
  <text x="40" y="69" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="end">10</text>

  <!-- Red Threshold Line at 9.0 -->
  <line x1="50" y1="80" x2="430" y2="80" stroke="#DC2626" stroke-width="2" stroke-dasharray="4,4"/>
  <text x="435" y="83" font-family="Inter, sans-serif" font-size="9" font-weight="700" fill="#DC2626">Порог: 9.0</text>

  <line x1="50" y1="125" x2="430" y2="125" stroke="#F1F5F9" stroke-width="1"/>
  <text x="40" y="129" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="end">5</text>

  <line x1="50" y1="170" x2="430" y2="170" stroke="#CBD5E1" stroke-width="1"/>
  <text x="40" y="174" font-family="Inter, sans-serif" font-size="9" fill="#94A3B8" text-anchor="end">0</text>

  <!-- Bar 1: Ivanov (11.0) -->
  <rect x="80" y="50" width="60" height="120" rx="4" fill="#15803D"/>
  <text x="110" y="42" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#15803D" text-anchor="middle">11.0</text>
  <text x="110" y="186" font-family="Inter, sans-serif" font-size="10" font-weight="600" fill="#0F172A" text-anchor="middle">Иванов (РОП)</text>

  <!-- Bar 2: Petrov (7.0 - Defect) -->
  <rect x="190" y="98" width="60" height="72" rx="4" fill="#DC2626"/>
  <text x="220" y="90" font-family="Inter, sans-serif" font-size="11" font-weight="800" fill="#DC2626" text-anchor="middle">7.0 (Брак)</text>
  <text x="220" y="186" font-family="Inter, sans-serif" font-size="10" font-weight="600" fill="#0F172A" text-anchor="middle">Петров (КАМ)</text>

  <!-- Bar 3: Sidorov (12.5) -->
  <rect x="300" y="32" width="60" height="138" rx="4" fill="#15803D"/>
  <text x="330" y="24" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#15803D" text-anchor="middle">12.5</text>
  <text x="330" y="186" font-family="Inter, sans-serif" font-size="10" font-weight="600" fill="#0F172A" text-anchor="middle">Сидоров (SDR)</text>
</svg>"""

with open("presentation/charts/speech_bar.svg", "w", encoding="utf-8") as f:
    f.write(speech_svg)

# 4. MONTE CARLO WATERFALL SVG
monte_carlo_svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 540 220" width="100%" height="100%">
  <!-- Target line at 5.0M -->
  <line x1="40" y1="40" x2="500" y2="40" stroke="#0F172A" stroke-width="1.5" stroke-dasharray="5,3"/>
  <text x="505" y="44" font-family="Inter, sans-serif" font-size="10" font-weight="800" fill="#0F172A">ПЛАН 5.0M ₽</text>

  <!-- Base line at 0 -->
  <line x1="40" y1="180" x2="500" y2="180" stroke="#CBD5E1" stroke-width="1"/>

  <!-- Step 1: As-Is (1.65M) -->
  <rect x="55" y="134" width="65" height="46" rx="3" fill="#64748B"/>
  <text x="87" y="126" font-family="Inter, sans-serif" font-size="10" font-weight="700" fill="#64748B" text-anchor="middle">1.65M</text>
  <text x="87" y="196" font-family="Inter, sans-serif" font-size="9.5" fill="#475569" text-anchor="middle">Факт As-Is</text>

  <!-- Step 2: Marketing (+0.5M) -->
  <rect x="145" y="120" width="65" height="14" rx="3" fill="#4F46E5"/>
  <text x="177" y="114" font-family="Inter, sans-serif" font-size="10" font-weight="700" fill="#4F46E5" text-anchor="middle">+0.50M</text>
  <text x="177" y="196" font-family="Inter, sans-serif" font-size="9.5" fill="#475569" text-anchor="middle">+30% Бюджет</text>

  <!-- Step 3: CR (+1.23M) -->
  <rect x="235" y="85" width="65" height="35" rx="3" fill="#4F46E5"/>
  <text x="267" y="79" font-family="Inter, sans-serif" font-size="10" font-weight="700" fill="#4F46E5" text-anchor="middle">+1.23M</text>
  <text x="267" y="196" font-family="Inter, sans-serif" font-size="9.5" fill="#475569" text-anchor="middle">+3% CR воронки</text>

  <!-- Step 4: KAM (+2.40M) -->
  <rect x="325" y="20" width="65" height="65" rx="3" fill="#6366F1"/>
  <text x="357" y="14" font-family="Inter, sans-serif" font-size="10" font-weight="700" fill="#4F46E5" text-anchor="middle">+2.40M</text>
  <text x="357" y="196" font-family="Inter, sans-serif" font-size="9.5" fill="#475569" text-anchor="middle">+1 КАМ (Штат)</text>

  <!-- Step 5: To-Be Total (5.775M) -->
  <rect x="425" y="20" width="75" height="160" rx="3" fill="#15803D"/>
  <text x="462" y="14" font-family="Inter, sans-serif" font-size="11" font-weight="800" fill="#15803D" text-anchor="middle">5.78M ₽</text>
  <text x="462" y="196" font-family="Inter, sans-serif" font-size="9.5" font-weight="700" fill="#15803D" text-anchor="middle">Прогноз To-Be</text>
</svg>"""

with open("presentation/charts/monte_carlo.svg", "w", encoding="utf-8") as f:
    f.write(monte_carlo_svg)

# 5. CASH FLOW CHART SVG
cash_flow_svg = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 500 200" width="100%" height="100%">
  <!-- Fact Cash-In -->
  <g transform="translate(10, 20)">
    <text x="140" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A" text-anchor="end">Факт Cash-In</text>
    <rect x="150" y="2" width="130" height="20" rx="3" fill="#15803D"/>
    <text x="290" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#15803D">1 650 000 ₽</text>
    <text x="390" y="16" font-family="Inter, sans-serif" font-size="9.5" fill="#64748B">На расчетном счете</text>
  </g>

  <!-- Expected 30d -->
  <g transform="translate(10, 60)">
    <text x="140" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A" text-anchor="end">Ожидаемый (30д)</text>
    <rect x="150" y="2" width="310" height="20" rx="3" fill="#B45309"/>
    <text x="470" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#B45309">4 075 000 ₽</text>
  </g>

  <!-- Overdue AR -->
  <g transform="translate(10, 100)">
    <text x="140" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A" text-anchor="end">Просрочка DSO</text>
    <rect x="150" y="2" width="25" height="20" rx="3" fill="#DC2626"/>
    <text x="185" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#DC2626">150 000 ₽</text>
    <text x="270" y="16" font-family="Inter, sans-serif" font-size="9.5" fill="#DC2626">DSO > 20 дней (INV-203)</text>
  </g>

  <!-- AI Loss Recovery -->
  <g transform="translate(10, 140)">
    <text x="140" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#0F172A" text-anchor="end">AI-Дожим отказников</text>
    <rect x="150" y="2" width="90" height="20" rx="3" fill="#4F46E5"/>
    <text x="250" y="16" font-family="Inter, sans-serif" font-size="11" font-weight="700" fill="#4F46E5">1 000 000 ₽</text>
    <text x="360" y="16" font-family="Inter, sans-serif" font-size="9.5" fill="#4F46E5">Возврат 440 000 ₽ (44%)</text>
  </g>
</svg>"""

with open("presentation/charts/cash_flow.svg", "w", encoding="utf-8") as f:
    f.write(cash_flow_svg)

print("Generated 5 vector SVG charts in presentation/charts/ successfully!")
