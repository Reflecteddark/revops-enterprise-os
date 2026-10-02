import pymupdf

svg_emblem = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="512" height="512">
  <defs>
    <!-- Background Gradient -->
    <radialGradient id="bg-grad" cx="50%" cy="45%" r="65%">
      <stop offset="0%" stop-color="#1e1b4b"/>
      <stop offset="60%" stop-color="#0f172a"/>
      <stop offset="100%" stop-color="#020617"/>
    </radialGradient>
    
    <!-- Accent Gradients -->
    <linearGradient id="chart-grad" x1="15%" y1="85%" x2="85%" y2="15%">
      <stop offset="0%" stop-color="#6366f1"/>
      <stop offset="45%" stop-color="#818cf8"/>
      <stop offset="85%" stop-color="#10b981"/>
      <stop offset="100%" stop-color="#34d399"/>
    </linearGradient>

    <linearGradient id="bar-grad-1" x1="0" y1="1" x2="0" y2="0">
      <stop offset="0%" stop-color="#312e81"/>
      <stop offset="100%" stop-color="#6366f1"/>
    </linearGradient>

    <linearGradient id="bar-grad-2" x1="0" y1="1" x2="0" y2="0">
      <stop offset="0%" stop-color="#3730a3"/>
      <stop offset="100%" stop-color="#818cf8"/>
    </linearGradient>

    <linearGradient id="bar-grad-3" x1="0" y1="1" x2="0" y2="0">
      <stop offset="0%" stop-color="#064e3b"/>
      <stop offset="100%" stop-color="#10b981"/>
    </linearGradient>

    <linearGradient id="ring-grad" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#6366f1" stop-opacity="0.8"/>
      <stop offset="50%" stop-color="#818cf8" stop-opacity="0.3"/>
      <stop offset="100%" stop-color="#10b981" stop-opacity="0.8"/>
    </linearGradient>

    <!-- Filters for Glow -->
    <filter id="emerald-glow" x="-30%" y="-30%" width="160%" height="160%">
      <feGaussianBlur stdDeviation="8" result="blur"/>
      <feMerge>
        <feMergeNode in="blur"/>
        <feMergeNode in="SourceGraphic"/>
      </feMerge>
    </filter>
    
    <filter id="indigo-glow" x="-30%" y="-30%" width="160%" height="160%">
      <feGaussianBlur stdDeviation="14" result="blur"/>
      <feMerge>
        <feMergeNode in="blur"/>
        <feMergeNode in="SourceGraphic"/>
      </feMerge>
    </filter>
  </defs>

  <!-- Base Circle Background -->
  <rect width="512" height="512" fill="#020617"/>
  <circle cx="256" cy="256" r="236" fill="url(#bg-grad)" stroke="url(#ring-grad)" stroke-width="8"/>
  <circle cx="256" cy="256" r="215" fill="none" stroke="#1e293b" stroke-width="2" stroke-dasharray="8 8" opacity="0.6"/>

  <!-- Audio Analysis Bars (Soundwave) -->
  <g opacity="0.95">
    <rect x="110" y="270" width="16" height="60" rx="8" fill="url(#bar-grad-1)"/>
    <rect x="142" y="230" width="16" height="110" rx="8" fill="url(#bar-grad-1)"/>
    <rect x="174" y="190" width="16" height="160" rx="8" fill="url(#bar-grad-2)"/>
    <rect x="206" y="220" width="16" height="120" rx="8" fill="url(#bar-grad-2)"/>
    <rect x="238" y="160" width="16" height="190" rx="8" fill="url(#bar-grad-2)"/>
    <rect x="270" y="210" width="16" height="135" rx="8" fill="url(#bar-grad-2)"/>
    <rect x="302" y="170" width="16" height="180" rx="8" fill="url(#bar-grad-3)"/>
    <rect x="334" y="210" width="16" height="125" rx="8" fill="url(#bar-grad-3)"/>
    <rect x="366" y="240" width="16" height="85" rx="8" fill="url(#bar-grad-3)"/>
    <rect x="398" y="275" width="16" height="50" rx="8" fill="url(#bar-grad-3)"/>
  </g>

  <!-- Breakthrough Sales Growth Line -->
  <path d="M 100 320 L 174 240 L 250 265 L 395 125" 
        fill="none" 
        stroke="url(#chart-grad)" 
        stroke-width="14" 
        stroke-linecap="round" 
        stroke-linejoin="round"
        filter="url(#indigo-glow)"/>

  <!-- Top Breakthrough Node / Arrow -->
  <circle cx="395" cy="125" r="22" fill="#10b981" filter="url(#emerald-glow)"/>
  <circle cx="395" cy="125" r="12" fill="#ffffff"/>
  
  <circle cx="250" cy="265" r="14" fill="#818cf8"/>
  <circle cx="250" cy="265" r="7" fill="#ffffff"/>

  <circle cx="174" cy="240" r="14" fill="#6366f1"/>
  <circle cx="174" cy="240" r="7" fill="#ffffff"/>

  <circle cx="100" cy="320" r="12" fill="#4f46e5"/>

  <!-- Bottom Brand Text -->
  <text x="256" y="425" text-anchor="middle" font-family="system-ui, -apple-system, sans-serif" font-weight="900" font-size="34" letter-spacing="4" fill="#f8fafc">
    AI-ROP
  </text>
  <text x="256" y="455" text-anchor="middle" font-family="system-ui, -apple-system, sans-serif" font-weight="600" font-size="14" letter-spacing="6" fill="#10b981">
    REVOPS INTELLIGENCE
  </text>
</svg>'''

# Version 2: Cute and Modern AI Bot Head with Soundwaves
svg_bot = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="512" height="512">
  <defs>
    <radialGradient id="bg-grad-bot" cx="50%" cy="40%" r="65%">
      <stop offset="0%" stop-color="#312e81"/>
      <stop offset="65%" stop-color="#0f172a"/>
      <stop offset="100%" stop-color="#020617"/>
    </radialGradient>
    <linearGradient id="bot-body" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#818cf8"/>
      <stop offset="50%" stop-color="#6366f1"/>
      <stop offset="100%" stop-color="#4338ca"/>
    </linearGradient>
    <linearGradient id="neon-cyan" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#34d399"/>
      <stop offset="100%" stop-color="#059669"/>
    </linearGradient>
    <linearGradient id="glow-ring" x1="0" y1="0" x2="1" y2="1">
      <stop offset="0%" stop-color="#10b981"/>
      <stop offset="100%" stop-color="#6366f1"/>
    </linearGradient>
    <filter id="bot-glow" x="-20%" y="-20%" width="140%" height="140%">
      <feGaussianBlur stdDeviation="10" result="blur"/>
      <feMerge>
        <feMergeNode in="blur"/>
        <feMergeNode in="SourceGraphic"/>
      </feMerge>
    </filter>
  </defs>

  <rect width="512" height="512" fill="#020617"/>
  <circle cx="256" cy="256" r="236" fill="url(#bg-grad-bot)" stroke="url(#glow-ring)" stroke-width="8"/>

  <!-- Antenna -->
  <line x1="256" y1="140" x2="256" y2="95" stroke="#818cf8" stroke-width="8" stroke-linecap="round"/>
  <circle cx="256" cy="90" r="16" fill="#10b981" filter="url(#bot-glow)"/>
  <circle cx="256" cy="90" r="7" fill="#ffffff"/>

  <!-- Bot Head -->
  <rect x="136" y="140" width="240" height="190" rx="48" fill="url(#bot-body)" stroke="#a5b4fc" stroke-width="4"/>
  
  <!-- Headphones on Sides -->
  <rect x="112" y="190" width="24" height="90" rx="12" fill="#10b981" filter="url(#bot-glow)"/>
  <rect x="376" y="190" width="24" height="90" rx="12" fill="#10b981" filter="url(#bot-glow)"/>

  <!-- Screen Face -->
  <rect x="160" y="165" width="192" height="140" rx="30" fill="#090d16" stroke="#1e293b" stroke-width="3"/>

  <!-- Glowing Eyes (Audio meters) -->
  <circle cx="205" cy="225" r="20" fill="#10b981" filter="url(#bot-glow)"/>
  <circle cx="205" cy="225" r="9" fill="#ffffff"/>
  
  <circle cx="307" cy="225" r="20" fill="#10b981" filter="url(#bot-glow)"/>
  <circle cx="307" cy="225" r="9" fill="#ffffff"/>

  <!-- Smile Wave -->
  <path d="M 215 270 Q 256 295 297 270" fill="none" stroke="#60a5fa" stroke-width="6" stroke-linecap="round"/>

  <!-- Sound Waves around bot -->
  <path d="M 90 205 Q 75 235 90 265" fill="none" stroke="#6366f1" stroke-width="5" stroke-linecap="round" opacity="0.8"/>
  <path d="M 422 205 Q 437 235 422 265" fill="none" stroke="#10b981" stroke-width="5" stroke-linecap="round" opacity="0.8"/>

  <!-- Bottom Brand Text -->
  <text x="256" y="415" text-anchor="middle" font-family="system-ui, -apple-system, sans-serif" font-weight="900" font-size="36" letter-spacing="4" fill="#f8fafc">
    AI-ROP
  </text>
  <text x="256" y="450" text-anchor="middle" font-family="system-ui, -apple-system, sans-serif" font-weight="600" font-size="14" letter-spacing="6" fill="#10b981">
    ГОЛОСОВОЙ ИИ-РОП
  </text>
</svg>'''

def render_svg(svg_content, out_path):
    doc = pymupdf.open(stream=svg_content.encode('utf-8'), filetype='svg')
    page = doc[0]
    # 512x512 with high quality
    mat = pymupdf.Matrix(2.0, 2.0) # 1024x1024 ultra sharp
    pix = page.get_pixmap(matrix=mat, alpha=False)
    pix.save(out_path)
    print(f"Saved {out_path}: {pix.width}x{pix.height}")

render_svg(svg_emblem, r"C:\Users\strel\Desktop\AI_ROP_Logo_Emblem.png")
render_svg(svg_bot, r"C:\Users\strel\Desktop\AI_ROP_Logo_Bot.png")
# Also create the primary default AI_ROP_Logo.png
render_svg(svg_emblem, r"C:\Users\strel\Desktop\AI_ROP_Logo.png")
