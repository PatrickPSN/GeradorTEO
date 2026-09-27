<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Gerador TEO · GACG</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js"></script>
<style>
  :root{
    /* marca (constante nos dois temas) */
    --blue:#0F766E; --blue-lt:#2DD4BF; --blue-dk:#115E59;
    --green:#F59E0B; --green-dk:#D97706;
    --r:18px; --r-sm:12px;
    /* ---- TEMA ESCURO (padrao) ---- */
    --ink:#f2f7f5; --ink-soft:#b7cbc7; --ink-mute:#7c9993;
    --glass:rgba(21,43,45,.62); --glass-2:rgba(15,31,34,.72);
    --stroke:rgba(92,190,176,.22); --stroke-2:rgba(92,190,176,.11);
    --ok:#2DD4BF; --warn:#F59E0B; --err:#FB7185;
    --glow-b:0 0 0 1px rgba(45,212,191,.24),0 14px 40px rgba(15,118,110,.30);
    --glow-g:0 0 0 1px rgba(245,158,11,.26),0 14px 40px rgba(217,119,6,.25);
    --page:#081416; --page-2:#0c2020;
    --au1:rgba(15,118,110,.30); --au2:rgba(245,158,11,.16); --au3:rgba(20,91,84,.28);
    --orb-op:.55; --grid-op:.05; --tbl-bg:rgba(8,17,32,.5); --tr-even:rgba(120,170,230,.03);
    --th-grad:linear-gradient(180deg,rgba(15,118,110,.78),rgba(17,94,89,.72)); --th-ink:#eafff9;
    --code-bg:rgba(45,212,191,.13); --logo-grad:linear-gradient(135deg,#d9fffa,#fef3c7);
    --h1-grad:linear-gradient(100deg,#f8fafc 10%,var(--blue-lt) 45%,#fbbf24 90%);
  }
  /* ---- TEMA CLARO ---- */
  body{
    --ink:#173b3a; --ink-soft:#456865; --ink-mute:#75918c;
    --glass:rgba(255,255,255,.78); --glass-2:rgba(255,255,255,.66);
    --stroke:rgba(15,118,110,.18); --stroke-2:rgba(15,118,110,.10);
    --ok:#0F766E; --warn:#B45309; --err:#E11D48;
    --glow-b:0 0 0 1px rgba(15,118,110,.18),0 14px 40px rgba(15,118,110,.15);
    --glow-g:0 0 0 1px rgba(217,119,6,.22),0 14px 40px rgba(217,119,6,.14);
    --page:#eef8f5; --page-2:#fff8e7;
    --au1:rgba(15,118,110,.17); --au2:rgba(245,158,11,.15); --au3:rgba(20,120,100,.13);
    --orb-op:.40; --grid-op:.05; --tbl-bg:rgba(255,255,255,.7); --tr-even:rgba(0,90,171,.035);
    --th-grad:linear-gradient(180deg,var(--blue),var(--blue-dk)); --th-ink:#fff;
    --code-bg:rgba(15,118,110,.10); --logo-grad:linear-gradient(135deg,var(--blue),var(--green));
    --h1-grad:linear-gradient(100deg,var(--blue-dk) 5%,var(--blue) 45%,var(--green-dk) 95%);
  }
  *{box-sizing:border-box}
  html{scroll-behavior:smooth}
  body{margin:0;font-family:"Manrope",system-ui,Segoe UI,Roboto,Arial,sans-serif;color:var(--ink);font-size:14px;
    background:var(--page);overflow-x:hidden;min-height:100vh;transition:color .4s,background .4s}
  code,.mono{font-family:"JetBrains Mono",ui-monospace,monospace}

  /* ===== fundo aurora animado ===== */
  .aurora{position:fixed;inset:0;z-index:-2;overflow:hidden;transition:background .4s;background:
    radial-gradient(1200px 700px at 12% -10%,var(--au1),transparent 60%),
    radial-gradient(1000px 680px at 100% 0%,var(--au2),transparent 55%),
    radial-gradient(900px 700px at 50% 120%,var(--au3),transparent 60%),  
    linear-gradient(180deg,var(--page) 0,var(--page-2) 100%)}
  .aurora:before,.aurora:after{content:"";position:absolute;width:60vmax;height:60vmax;border-radius:50%;
    filter:blur(80px);opacity:var(--orb-op);mix-blend-mode:screen;animation:float 22s ease-in-out infinite}
  .aurora:before{background:radial-gradient(circle,#0072CE,transparent 65%);top:-15vmax;left:-10vmax}
  .aurora:after{background:radial-gradient(circle,#00E08A,transparent 65%);bottom:-18vmax;right:-12vmax;animation-delay:-8s}
  @keyframes float{0%,100%{transform:translate(0,0) scale(1)}50%{transform:translate(6vmax,4vmax) scale(1.12)}}
  /* grade sutil */
  .grid-bg{position:fixed;inset:0;z-index:-1;pointer-events:none;
    background-image:linear-gradient(rgba(120,170,230,var(--grid-op)) 1px,transparent 1px),linear-gradient(90deg,rgba(120,170,230,var(--grid-op)) 1px,transparent 1px);
    background-size:46px 46px;mask-image:radial-gradient(circle at 50% 30%,#000 0,transparent 78%)}

  /* ===== header ===== */
  header{position:relative;padding:30px 28px 22px;border-bottom:1px solid var(--stroke)}
  .head-inner{max-width:1200px;margin:0 auto}
  .brand{display:flex;align-items:center;gap:18px;flex-wrap:wrap}
  .logo{
    display:flex;
    align-items:center;
    justify-content:center;

    background:none;
    box-shadow:none;

    padding:0;
    min-width:auto;
    height:auto;
}
  .logo span{font-weight:800;font-size:27px;letter-spacing:.5px;line-height:1;transition:.4s;
    background:linear-gradient(120deg,var(--blue-dk),var(--green-dk));-webkit-background-clip:text;background-clip:text;-webkit-text-fill-color:transparent}
  .sabesp-logo{display:flex;align-items:center;gap:8px;color:#087ea5!important;-webkit-text-fill-color:#087ea5!important;background:none!important;font-size:22px!important;letter-spacing:0!important}
  .sabesp-symbol{position:relative;width:32px;height:38px;display:inline-block;border-radius:4px 17px 17px 4px;background:linear-gradient(135deg,#16b9d8 0 42%,#087ea5 43% 62%,#24c5df 63%);box-shadow:0 4px 12px rgba(8,126,165,.25)}
  .sabesp-symbol:after{content:"";position:absolute;left:6px;right:3px;bottom:8px;height:7px;border-radius:0 0 12px 12px;background:#fff;opacity:.9}
  .sabesp-header-logo{
    width:80px;
    height:auto;
    object-fit:contain;
}
  body.light .logo .sabesp-logo{-webkit-text-fill-color:#087ea5!important;color:#087ea5!important}
  body.light .logo .sabesp-symbol{background:linear-gradient(135deg,#16b9d8 0 42%,#087ea5 43% 62%,#24c5df 63%)!important;-webkit-text-fill-color:initial}
  .titlewrap{flex:1}
  .titlewrap h1{
    margin:0;
    font-size:30px;
    font-weight:800;
    letter-spacing:-.3px;
    line-height:1.05;

    color:#0F766E;
}

  /* ===== toggle de tema ===== */
  .theme-toggle{margin-left:auto;display:inline-flex;align-items:center;gap:0;background:var(--glass);border:1px solid var(--stroke);
    border-radius:30px;padding:4px;cursor:pointer;-webkit-backdrop-filter:blur(10px);backdrop-filter:blur(10px);position:relative;transition:.3s;flex-shrink:0}
  .theme-toggle:hover{border-color:var(--blue-lt)}
  .theme-toggle .opt{width:34px;height:30px;border-radius:24px;display:flex;align-items:center;justify-content:center;font-size:15px;
    position:relative;z-index:1;color:var(--ink-mute);transition:.3s}
  .theme-toggle .opt.on{color:#fff}
  .theme-toggle .knob{position:absolute;top:4px;left:4px;width:34px;height:30px;border-radius:24px;
    background:linear-gradient(135deg,var(--blue),var(--green-dk));box-shadow:0 4px 12px rgba(0,114,206,.5);transition:transform .35s cubic-bezier(.4,1.4,.5,1)}
  body.light .theme-toggle .knob{transform:translateX(34px);background:linear-gradient(135deg,var(--warn),#ffd98a)}
  @keyframes shine{to{background-position:220% center}}
  .titlewrap .sub{margin-top:4px;color:var(--ink-soft);font-size:13.5px;font-weight:500;display:flex;align-items:center;gap:8px}
  .pill{display:inline-flex;align-items:center;gap:6px;padding:3px 10px;border-radius:30px;font-size:11px;font-weight:700;
    background:rgba(0,224,138,.12);color:var(--green);border:1px solid rgba(0,224,138,.3)}
  .dot{width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 0 0 rgba(0,224,138,.7);animation:pulse 2s infinite}
  @keyframes pulse{0%{box-shadow:0 0 0 0 rgba(0,224,138,.6)}70%{box-shadow:0 0 0 8px rgba(0,224,138,0)}100%{box-shadow:0 0 0 0 rgba(0,224,138,0)}}

  .wrap{max-width:1200px;margin:0 auto;padding:26px 28px 90px}

  /* ===== glass card base ===== */
  .glass{background:linear-gradient(180deg,var(--glass),var(--glass-2));border:1px solid var(--stroke);
    border-radius:var(--r);box-shadow:0 20px 50px rgba(0,0,0,.35),inset 0 1px 0 rgba(255,255,255,.05);
    -webkit-backdrop-filter:blur(14px);backdrop-filter:blur(14px);transition:background .4s,border-color .4s}
  body.light .glass{box-shadow:0 16px 42px rgba(60,90,130,.14),inset 0 1px 0 rgba(255,255,255,.7)}

  .instr{padding:20px 22px;margin-bottom:22px;position:relative;overflow:hidden}
  .instr:before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:linear-gradient(var(--blue),var(--green))}
  .instr h2{margin:0 0 9px;font-size:16px;font-weight:800;color:var(--ink);display:flex;align-items:center;gap:10px}
  .instr h2 .ico{width:26px;height:26px;border-radius:8px;display:inline-flex;align-items:center;justify-content:center;font-size:14px;font-weight:800;
    background:linear-gradient(135deg,var(--blue),var(--green));color:#04101f;box-shadow:0 6px 18px rgba(0,224,138,.35)}
  .instr p{margin:0;color:var(--ink-soft);font-size:13px;line-height:1.65}

  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:18px}
  @media(max-width:860px){.grid2{grid-template-columns:1fr}}
  .card{padding:20px}
  .card h3{margin:0 0 4px;font-size:14px;font-weight:700;color:var(--ink);display:flex;align-items:center;gap:9px}
  .card h3 .n{width:24px;height:24px;border-radius:7px;display:inline-flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;
    background:rgba(0,114,206,.18);color:var(--blue);border:1px solid rgba(0,114,206,.35)}
  .card .hint{color:var(--ink-mute);font-size:12px;margin:0 0 14px;line-height:1.5}
  .teo-form{margin:22px 0}
  .teo-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:13px 16px}
  .teo-grid label,.teo-grid fieldset{color:var(--ink-soft);font-size:12px;font-weight:700;min-width:0}
  .teo-grid label,.teo-grid legend{color:#111}
  .teo-grid label>input[type=text],.teo-grid label>input[type=date],.teo-grid textarea{display:block;width:100%;margin-top:6px;padding:10px 11px;border:1px solid var(--stroke);border-radius:10px;background:var(--tbl-bg);color:#111;font:500 13px inherit;resize:vertical}
  .teo-grid label>input:focus,.teo-grid textarea:focus{outline:none;border-color:var(--blue-lt);box-shadow:0 0 0 3px rgba(0,114,206,.2)}
  .teo-grid .wide{grid-column:1/-1}
  .teo-grid fieldset{border:1px solid var(--stroke);border-radius:10px;padding:8px 11px;display:flex;align-items:center;gap:14px;margin:0}
  .teo-grid legend{padding:0 5px;color:var(--ink-mute);font-size:11px}
  .teo-grid .inline,.teo-grid .check-field{display:flex;align-items:center;gap:7px;font-weight:600}
  .teo-grid .check-field{align-self:end;padding:10px 0}
  .teo-grid input[type=checkbox],.teo-grid input[type=radio]{accent-color:var(--blue)}
  .teo-grid input[type=file]{display:none}
  .file-button{display:inline-flex;margin-top:6px;padding:9px 12px;border-radius:9px;background:rgba(0,114,206,.15);border:1px solid var(--stroke);color:var(--blue-lt);cursor:pointer}
  .file-name{margin-left:9px;color:var(--ink-mute);font-weight:500;word-break:break-word}
  .teo-actions{margin-bottom:0}
  @media(max-width:700px){.teo-grid{grid-template-columns:1fr}.teo-grid .wide{grid-column:auto}.teo-grid fieldset{min-height:44px}}

  /* ===== dropzone ===== */
  .drop{position:relative;border:1.5px dashed rgba(120,170,230,.3);border-radius:var(--r-sm);padding:30px 18px;text-align:center;cursor:pointer;
    transition:.22s cubic-bezier(.2,.8,.2,1);background:var(--tbl-bg);overflow:hidden}
  .drop:before{
    display:none;
}
  .drop:hover{border-color:var(--blue-lt);transform:translateY(-2px);box-shadow:var(--glow-b)}
  .drop:hover:before,.drop.drag:before{opacity:1}
  .drop.drag{border-color:var(--blue-lt);background:rgba(0,114,206,.10)}
  .drop .ic{width:46px;height:46px;margin:0 auto 10px;border-radius:14px;display:flex;align-items:center;justify-content:center;
    background:linear-gradient(135deg,rgba(0,114,206,.25),rgba(0,224,138,.18));border:1px solid var(--stroke);font-size:22px}
  .drop .big{font-size:13.5px;font-weight:700;color:var(--ink);position:relative}
  .drop .sub{color:var(--ink-mute);font-size:12px;margin-top:4px;position:relative}
  .drop.loaded{border-color:var(--green);border-style:solid;background:rgba(0,224,138,.08);box-shadow:var(--glow-g)}
  .drop.loaded .ic{background:linear-gradient(135deg,var(--green),var(--green-dk));color:#04101f}
  .filemeta{font-size:12px;color:var(--green);margin-top:10px;word-break:break-all;font-weight:600;position:relative}
  input[type=file]{display:none}

  /* ===== opções ===== */
  .opts{display:flex;flex-wrap:wrap;gap:16px;align-items:center;margin:22px 0 4px}
  .opts label{display:flex;gap:9px;align-items:center;color:var(--ink-soft);font-size:13px;font-weight:600}
  select{font-family:inherit;background:var(--tbl-bg);color:var(--ink);border:1px solid var(--stroke);border-radius:10px;padding:9px 12px;font-size:13px;font-weight:500;
    appearance:none;cursor:pointer;background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='12' viewBox='0 0 24 24' fill='none' stroke='%233aa0ff' stroke-width='3'%3E%3Cpath d='M6 9l6 6 6-6'/%3E%3C/svg%3E");background-repeat:no-repeat;background-position:right 11px center;padding-right:32px;transition:.15s}
  select:hover{border-color:var(--blue-lt)}
  select:focus{outline:none;border-color:var(--blue-lt);box-shadow:0 0 0 3px rgba(0,114,206,.2)}

  /* ===== botões ===== */
  .actions{display:flex;gap:13px;margin:24px 0 6px;align-items:center;flex-wrap:wrap}
  button{font-family:inherit;font-size:14px;font-weight:700;border:0;border-radius:12px;padding:13px 26px;cursor:pointer;
    transition:.2s cubic-bezier(.2,.8,.2,1);position:relative;overflow:hidden}
  button:active{transform:scale(.97)}
  .btn{
    background:#0F766E;
    color:#fff;
    box-shadow:0 3px 8px rgba(0,0,0,.12);
}
  .btn:before{content:"";position:absolute;top:0;left:-120%;width:60%;height:100%;background:linear-gradient(110deg,transparent,rgba(255,255,255,.4),transparent);transition:.5s}
  .btn:hover:not(:disabled){
    transform:translateY(-1px);
    box-shadow:0 6px 12px rgba(0,0,0,.18);
}
  .btn:hover:not(:disabled):before{left:130%}
  .btn:disabled{background:rgba(80,100,130,.35);color:var(--ink-mute);cursor:not-allowed;box-shadow:none}
  .btn.ghost{background:rgba(120,170,230,.08);border:1px solid var(--stroke);color:var(--ink);box-shadow:none}
  .btn.ghost:hover{background:rgba(120,170,230,.16);border-color:var(--blue-lt)}
  .btn.ok{
    background:#0F766E;
    color:#fff;
    box-shadow:0 3px 8px rgba(0,0,0,.12);
    border:1px solid rgba(0,0,0,.08);
}

.btn.ok:hover{
    transform:translateY(-1px);
    box-shadow:0 6px 12px rgba(0,0,0,.18);
}

  /* ===== mensagens ===== */
  .msg{padding:12px 16px;border-radius:12px;margin:16px 0;font-size:13px;display:none;font-weight:500}
  .msg.show{display:block;animation:slideIn .35s ease}
  .msg.error{
    background:#FEE2E2;
    color:#991B1B;
    border:1px solid #FCA5A5;
}
  @keyframes slideIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:none}}

  /* ===== resultados ===== */
  #results{margin-top:8px}
  #results.reveal{animation:reveal .5s ease}
  @keyframes reveal{from{opacity:0;transform:translateY(14px)}to{opacity:1;transform:none}}
  .stat-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin:10px 0 20px}
  @media(max-width:860px){.stat-row{grid-template-columns:repeat(2,1fr)}}
  .stat{padding:18px 18px 16px;position:relative;overflow:hidden}
  .stat:before{content:"";position:absolute;top:0;left:0;right:0;height:3px;background:linear-gradient(90deg,var(--blue),var(--blue-lt))}
  .stat .v{font-size:27px;font-weight:800;letter-spacing:-.5px;color:var(--ink);line-height:1.1}
  .stat .l{font-size:11.5px;color:var(--ink-mute);margin-top:4px;text-transform:uppercase;letter-spacing:.6px;font-weight:600}
  .stat .spark{position:absolute;right:14px;top:14px;width:34px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:16px;opacity:.85}
  .stat.ok:before{background:linear-gradient(90deg,var(--green),#7cffc4)} .stat.ok .v{color:var(--green)} .stat.ok .spark{background:rgba(0,224,138,.14)}
  .stat.warn:before{background:linear-gradient(90deg,var(--warn),#ffe09a)} .stat.warn .v{color:var(--warn)} .stat.warn .spark{background:rgba(255,196,77,.14)}
  .stat.err:before{background:linear-gradient(90deg,var(--err),#ffa7b4)} .stat.err .v{color:var(--err)} .stat.err .spark{background:rgba(255,107,129,.14)}
  .stat .spark.b{background:rgba(0,114,206,.16)}

  /* ===== abas ===== */
  .tabs{display:flex;gap:6px;margin:20px 0 0;flex-wrap:wrap}
  .tab{padding:10px 16px;cursor:pointer;color:var(--ink-soft);font-size:13px;font-weight:600;border-radius:11px;border:1px solid transparent;transition:.18s}
  .tab:hover{background:rgba(120,170,230,.08);color:var(--ink)}
  .tab.active{color:var(--ink);background:linear-gradient(135deg,rgba(0,114,206,.3),rgba(0,224,138,.18));border-color:var(--stroke);box-shadow:inset 0 1px 0 rgba(255,255,255,.08)}
  body.light .tab.active{background:linear-gradient(135deg,rgba(0,114,206,.16),rgba(0,168,89,.12))}
  .panel-body{display:none;padding-top:16px;animation:fadeUp .35s ease}
  .panel-body.active{display:block}
  @keyframes fadeUp{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}

  /* ===== tabelas ===== */
  .tbl-scroll{max-height:460px;overflow:auto;border:1px solid var(--stroke);border-radius:var(--r-sm);background:var(--tbl-bg)}
  table{width:100%;border-collapse:collapse;font-size:12.5px;color:#111}
  th,td{padding:9px 12px;text-align:left;border-bottom:1px solid var(--stroke-2)}
  th{color:var(--th-ink);font-weight:700;position:sticky;top:0;font-size:11px;text-transform:uppercase;letter-spacing:.5px;
    background:var(--th-grad);-webkit-backdrop-filter:blur(6px);backdrop-filter:blur(6px)}
  tbody tr{transition:.12s}
  tbody tr:nth-child(even){background:var(--tr-even)}
  tbody tr:hover{background:rgba(0,114,206,.12)}
  td.num,th.num{text-align:right;font-variant-numeric:tabular-nums;font-family:"JetBrains Mono",monospace}
  td:first-child{font-family:"JetBrains Mono",monospace;color:var(--ink-soft)}
  .badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:30px;font-size:11px;font-weight:700}
  .b-ok{background:rgba(0,224,138,.14);color:var(--green)} .b-warn{background:rgba(255,196,77,.14);color:var(--warn)} .b-err{background:rgba(255,107,129,.14);color:var(--err)}
  .alert{padding:12px 15px;border-radius:12px;margin-bottom:9px;font-size:13px;border-left:4px solid;font-weight:500}
  .alert.warn{background:rgba(255,196,77,.08);border-color:var(--warn);color:#ffe0a3} .alert.err{background:rgba(255,107,129,.08);border-color:var(--err);color:#ffb3bf} .alert.ok{background:rgba(0,224,138,.08);border-color:var(--green);color:#9bffd2}

  .hidden{display:none}
  code{background:var(--code-bg);color:var(--blue);padding:1.5px 7px;border-radius:6px;font-size:12px;border:1px solid rgba(0,114,206,.18)}
  .neg{color:var(--err)}
  /* scrollbar */
  ::-webkit-scrollbar{width:10px;height:10px}
  ::-webkit-scrollbar-track{background:transparent}
  ::-webkit-scrollbar-thumb{background:rgba(120,170,230,.25);border-radius:10px;border:2px solid transparent;background-clip:content-box}
  ::-webkit-scrollbar-thumb:hover{background:rgba(120,170,230,.4);background-clip:content-box}
  .foot{text-align:center;color:var(--ink-mute);font-size:11.5px;margin-top:30px;letter-spacing:.3px}
</style>
</head>
<body>
<body class="light">
<div class="aurora"></div>
<div class="grid-bg"></div>

<header>
  <div class="head-inner">
    <div class="brand">
      <div class="logo" aria-label="Logo SABESP"><img id="sabesp-header-image" class="sabesp-header-logo" alt="SABESP"></div>
      <div class="titlewrap">
        <h1>Gerador TEO</h1>
        <div class="sub">Termo de Entrada em Opera&ccedil;&atilde;o <span class="pill"><span class="dot"></span>GACG</span></div>
      </div>

    </div>
  </div>
</header>

<div class="wrap">
  <div class="instr glass">
    <h2><span class="ico">i</span> Instru&ccedil;&otilde;es</h2>
    <p>Preencha os dados do termo, anexe a logo da contratada e carregue a planilha com os ativos do Anexo I. A tabela será exibida para conferência e incorporada ao PDF final.</p>
  </div>

  <div class="card glass teo-form">
    <h3><span class="n">A</span> Dados do TEO</h3>
    <p class="hint">Preencha os dados que serao usados no formulario. O Anexo I sera montado a partir da planilha processada.</p>
    <div class="teo-grid">
      <label>Elemento PEP<input id="teo-pep" type="text"></label>
      <label>Nº. Contrato<input id="teo-contrato" type="text"></label>
      <label class="wide">Descrição do PEP<input id="teo-descricao" type="text"></label>
      <label>Administrador do Contrato<input id="teo-administrador" type="text"></label>
      <label>Empresa Contratada<input id="teo-empresa" type="text"></label>
      <label class="wide">Escopo do Contrato<textarea id="teo-escopo" rows="3"></textarea></label>
      <label class="wide">Endereço<input id="teo-endereco" type="text"></label>
      <label>Município<input id="teo-municipio" type="text"></label>
      <label>Bairro<input id="teo-bairro" type="text"></label>
      <label>Data Entrada em Operação<input id="teo-data-operacao" type="date"></label>
      <label class="check-field"><input id="teo-varias-datas" type="checkbox"> Varias datas - Vide anexo I</label>
      <fieldset><legend>Execução</legend><label class="inline"><input type="radio" name="teo-execucao" value="Parcial"> Parcial</label><label class="inline"><input type="radio" name="teo-execucao" value="Total"> Total</label></fieldset>
      <fieldset><legend>Houve desativação?</legend><label class="inline"><input type="radio" name="teo-desativacao" value="Sim"> Sim</label><label class="inline"><input type="radio" name="teo-desativacao" value="Nao"> Nao</label></fieldset>
      <label>Responsavel Sabesp - Nome<input id="teo-sabesp-nome" type="text"></label>
      <label>Responsavel Sabesp - Matricula<input id="teo-sabesp-matricula" type="text"></label>
      <label>Responsavel Contratada - Nome<input id="teo-contratada-nome" type="text"></label>
      <label>Data da Inspecao<input id="teo-inspecao" type="date"></label>
      <label class="wide">Logo da Contratada <input id="teo-logo" type="file" accept="image/png,image/jpeg"><span class="file-button">Selecionar imagem</span><span id="teo-logo-name" class="file-name">Nenhuma imagem selecionada</span></label>
    </div>
    <div class="actions teo-actions">
      <button class="btn ok" id="btn-pdf" disabled>&#128196; Gerar TEO em PDF</button>
    </div>
  </div>

  <div class="card glass upload-card">
    <h3><span class="n">1</span> Planilha Anexo 2 ou Anexo 3</h3>
    <p class="hint">Anexe apenas uma planilha. O cabeçalho deve estar na linha 7 e os dados começam na linha 8.</p>
    <div class="drop" id="drop-anexo">
      <div class="ic">
    <svg width="24" height="24" viewBox="0 0 24 24"
         fill="none" stroke="currentColor"
         stroke-width="2">
        <rect x="3" y="3" width="18" height="18" rx="2"/>
        <line x1="3" y1="9" x2="21" y2="9"/>
        <line x1="9" y1="3" x2="9" y2="21"/>
        <line x1="15" y1="3" x2="15" y2="21"/>
    </svg>
</div>
      <div class="big">Clique ou arraste o Anexo 2 ou 3</div>
      <div class="sub">.xlsx / .xls / .xlsm</div>
      <div class="filemeta hidden" id="meta-anexo"></div>
    </div>
    <input type="file" id="file-anexo" accept=".xlsx,.xls,.xlsm">
  </div>

  <div class="actions">
    <button class="btn" id="btn-run" disabled>&#9889; Montar Anexo I</button>
    <button class="btn ghost" id="btn-reset">Limpar</button>
  </div>

  <div class="msg" id="msg"></div>

  <div id="results" class="hidden">
    <div class="card glass preview-card">
      <h3><span class="n">2</span> Pré-visualização do Anexo I</h3>
      <div class="tbl-scroll"><table id="t-prev"></table></div>
    </div>
  </div>

  <div class="foot">GACG &ndash; Governan&ccedil;a Regulat&oacute;ria de Ativos</div>
</div>

<script>
var EPS = 0.005;
var state = { anexo:null, result:null, logoData:null, logoType:null, sabespLogoData:null };

/* ---------- utilidades ---------- */
function toNum(v){
  if(v==null||v==="") return 0;
  if(typeof v==="number") return v;
  var s=String(v).trim().replace(/\s/g,"").replace(/R\$/i,"");
  if(s.indexOf(",")>-1 && s.indexOf(".")>-1){ s=s.replace(/\./g,"").replace(",","."); }
  else if(s.indexOf(",")>-1){ s=s.replace(",","."); }
  var n=parseFloat(s);
  return isNaN(n)?0:n;
}
function fmtBRL(n){ return (n||0).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2}); }
function esc(s){
  if(s==null) return "";
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
}

function sheetRows(ws,headerIndex){
  var arr = XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
  var start=headerIndex||0;
  if(!arr.length||!arr[start]) return {hdr:[],rows:[]};
  var hdr = arr[start].map(function(h){ return h==null?"":String(h).trim(); });
  var nmap = hdr.map(function(h){ return String(h).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]/gi,"").toUpperCase(); });
  var rows = [];
  for(var i=start+1;i<arr.length;i++){
    var r=arr[i];
    if(!r) continue;
    var empty=true; for(var k=0;k<r.length;k++){ if(r[k]!=null && r[k]!=="") { empty=false; break; } }
    if(empty) continue;
    var o={}; o.__n={};
    for(var j=0;j<hdr.length;j++){ o[hdr[j]]=r[j]; o.__n[nmap[j]]=r[j]; }
    rows.push(o);
  }
  return {hdr:hdr, nmap:nmap, rows:rows};
}

/* ---------- leitura de arquivo ---------- */
function readWorkbook(file){
  return new Promise(function(res,rej){
    var fr=new FileReader();
    fr.onload=function(e){ try{ res(XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true,cellStyles:true,cellNF:true})); }catch(err){ rej(err); } };
    fr.onerror=rej;
    fr.readAsArrayBuffer(file);
  });
}

/* ---------- nucleo do calculo antigo removido do fluxo TEO ---------- */
/*
function processar(){
  var method = document.getElementById("opt-method").value;
  var negMode = document.getElementById("opt-neg").value;

  var wsSaldo = findSheet(state.log,["chaves com saldo","chave saldo","saldo"]);
  var wsModelo = findSheet(state.log,["modelo upload","modelo","upload"]);
  if(!wsSaldo) throw new Error("Nao encontrei a guia 'Chaves com Saldo' no LOG.");
  if(!wsModelo) throw new Error("Nao encontrei a guia 'Modelo Upload' no LOG.");
  var auxWs = findSheet(state.aux,["planilha1","base","aux"]) || state.aux.Sheets[state.aux.SheetNames[0]];

  var saldo = sheetRows(wsSaldo);
  var modelo = sheetRows(wsModelo);
  var aux = sheetRows(auxWs);

  // 1) Oferta: agrupa chaves por (pep|classe)
  var grp = {};
  var chaveMeta = {};
  for(var i=0;i<saldo.rows.length;i++){
    var rr=saldo.rows[i];
    var chave = pick(rr,"CHAVE"); if(!chave) continue;
    var parts = String(chave).split("|");
    var pep=parts[0], classe=(parts[1]||"").trim(), mes=parts[3]||"";
    var s = toNum(pick(rr,"SALDO_DISPONIVEL","SALDO DISPONIVEL","SALDO"));
    var g = pep+"||"+classe;
    if(!grp[g]) grp[g]=[];
    grp[g].push({chave:chave,mes:mes,saldo:s});
    chaveMeta[chave]={pep:pep,classe:classe,mes:mes,saldo:s};
  }

  // 2) Demanda: base auxiliar (ativo+classe)
  var demand={};
  var assetMeta={};
  var order=[];
  for(var a=0;a<aux.rows.length;a++){
    var ar=aux.rows[a];
    var ativo = pick(ar,"ATIVO","IMOBILIZADO"); if(ativo==null) continue;
    var classeA = String(pick(ar,"CLASSE")||"").trim();
    var val = pick(ar,"VALOR");
    if(!(ativo in assetMeta)){
      assetMeta[ativo]={
        inv:pick(ar,"N INVENTARIO","NO INVENTARIO","NUMERO DO INVENTARIO","INVENTARIO"),
        data:pick(ar,"DATA UNITIZACAO","DATA DA UNITIZACAO","DATA"),
        gear:pick(ar,"N SOLICITACAO GEAR","NO SOLICITACAO GEAR","SOLICITACAO GEAR","GEAR")
      };
      order.push(ativo);
    }
    if(val==null||val==="") continue;
    var dk=ativo+"||"+classeA;
    demand[dk]=(demand[dk]||0)+toNum(val);
  }

  // 3) Alocacao
  var alloc={};
  var reconChave={};
  var reconAC=[];
  var alerts=[];
  for(var gk in grp){
    var gp=gk.split("||"); var gpep=gp[0], gclasse=gp[1];
    var lst=grp[gk].slice().sort(function(x,y){ return x.mes<y.mes?-1:(x.mes>y.mes?1:0); });
    var net=0; for(var n0=0;n0<lst.length;n0++) net+=lst[n0].saldo;
    var sumPos=0; for(var n1=0;n1<lst.length;n1++){ if(lst[n1].saldo>0) sumPos+=lst[n1].saldo; }
    var pos=[]; for(var n2=0;n2<lst.length;n2++){ if(lst[n2].saldo>0) pos.push({chave:lst[n2].chave,mes:lst[n2].mes,rem:lst[n2].saldo,saldo:lst[n2].saldo}); }
    var negCount=0; for(var n3=0;n3<lst.length;n3++){ if(lst[n3].saldo<0) negCount++; }
    var cap = (negMode==="net")?net:sumPos;

    var dems=[];
    for(var od=0;od<order.length;od++){
      var at=order[od]; var nd=demand[at+"||"+gclasse]||0;
      if(nd>EPS) dems.push({ativo:at,need:nd});
    }
    if(dems.length===0) continue;
    var totalDem=0; for(var td=0;td<dems.length;td++) totalDem+=dems[td].need;

    if(totalDem > cap+EPS)
      alerts.push({lvl:"err",msg:"INVIAVEL - "+gpep+" / classe "+gclasse+": demanda R$ "+fmtBRL(totalDem)+" maior que a capacidade "+(negMode==="net"?"liquida":"positiva")+" R$ "+fmtBRL(cap)+"."});
    if(negCount)
      alerts.push({lvl:"warn",msg:gpep+" / classe "+gclasse+": "+negCount+" chave(s) com saldo negativo ("+(negMode==="net"?"compensadas no liquido":"ignoradas")+"); recebem 0."});

    for(var d=0;d<dems.length;d++){
      var rest=dems[d].need;
      if(method==="prop" && sumPos>EPS){
        // 1a passada proporcional ao saldo, mas NUNCA estourando o teto de cada chave (col. J)
        for(var p=0;p<pos.length;p++){
          if(rest<=EPS) break;
          var share=dems[d].need*(pos[p].saldo/sumPos);
          var vv=Math.min(share,pos[p].rem,rest);
          if(vv>EPS){ var kp=pos[p].chave+"||"+dems[d].ativo; alloc[kp]=(alloc[kp]||0)+vv; reconChave[pos[p].chave]=(reconChave[pos[p].chave]||0)+vv; pos[p].rem-=vv; rest-=vv; }
        }
        // mop-up FIFO do residuo de arredondamento proporcional nas chaves que ainda tem saldo
        for(var p2=0;p2<pos.length;p2++){
          if(rest<=EPS) break;
          var take2=Math.min(rest,pos[p2].rem);
          if(take2>EPS){ var kp2=pos[p2].chave+"||"+dems[d].ativo; alloc[kp2]=(alloc[kp2]||0)+take2; reconChave[pos[p2].chave]=(reconChave[pos[p2].chave]||0)+take2; pos[p2].rem-=take2; rest-=take2; }
        }
      } else {
        for(var q=0;q<pos.length;q++){
          if(rest<=EPS) break;
          var take=Math.min(rest,pos[q].rem);
          if(take>EPS){ var kf=pos[q].chave+"||"+dems[d].ativo; alloc[kf]=(alloc[kf]||0)+take; reconChave[pos[q].chave]=(reconChave[pos[q].chave]||0)+take; pos[q].rem-=take; rest-=take; }
        }
      }
      reconAC.push({ativo:dems[d].ativo,classe:gclasse,demanda:dems[d].need,alocado:dems[d].need-Math.max(rest,0)});
      if(rest>EPS) alerts.push({lvl:"err",msg:"FALTA - ativo "+dems[d].ativo+" / classe "+gclasse+": R$ "+fmtBRL(rest)+" nao alocado (sem saldo)."});
    }
  }

  // 4) Preenche Modelo Upload (mantem as linhas do template)
  var out=[];
  for(var m=0;m<modelo.rows.length;m++){
    var mr=modelo.rows[m];
    var mAtivo=pick(mr,"IMOBILIZADO","ATIVO");
    var mChave=pick(mr,"CHAVE");
    var mVal=alloc[mChave+"||"+mAtivo]||0;
    var meta=assetMeta[mAtivo]||{};
    out.push({
      IMOBILIZADO:mAtivo, CHAVE:mChave,
      DESCRICAO_CLASSE_CUSTO:pick(mr,"DESCRICAO CLASSE CUSTO","DESCRICAO_CLASSE_CUSTO"),
      DESCRICAO_MATERIAL:pick(mr,"DESCRICAO MATERIAL","DESCRICAO_MATERIAL"),
      VALOR_CONSUMIDO:Math.round(mVal*100)/100,
      INVENTARIO:(meta.inv!=null?meta.inv:null), DATA:(meta.data!=null?meta.data:null), GEAR:(meta.gear!=null?meta.gear:null)
    });
  }

  // metricas
  var ativosSet={}; for(var z=0;z<out.length;z++) ativosSet[out[z].IMOBILIZADO]=1;
  var ativos=Object.keys(ativosSet).length;
  var totalAloc=0; for(var z2=0;z2<out.length;z2++) totalAloc+=out[z2].VALOR_CONSUMIDO;
  var totDem=0,totAl=0; for(var z3=0;z3<reconAC.length;z3++){ totDem+=reconAC[z3].demanda; totAl+=reconAC[z3].alocado; }
  var cob= totDem>0 ? (totAl/totDem*100) : 100;

  // localiza o NOME da guia Modelo Upload no workbook original (para reescrever in-place)
  var modeloName=null;
  for(var sn=0;sn<state.log.SheetNames.length;sn++){ if(state.log.Sheets[state.log.SheetNames[sn]]===wsModelo){ modeloName=state.log.SheetNames[sn]; break; } }

  state.result={out:out,reconAC:reconAC,chaveMeta:chaveMeta,reconChave:reconChave,alerts:alerts,ativos:ativos,totalAloc:totalAloc,cob:cob,
                alloc:alloc,assetMeta:assetMeta,modeloName:modeloName};
  render();
}

// ---------- render antigo ----------
function fmtCell(v){ if(v==null||v==="") return "-"; return v; }
function fmtDate(v){ if(v==null||v==="") return "-"; if(v instanceof Date) return v.toLocaleDateString("pt-BR"); return v; }
function render(){
  var r=state.result;
  document.getElementById("results").classList.remove("hidden");
  document.getElementById("results").classList.add("reveal");
  document.getElementById("btn-download").classList.remove("hidden");
  document.getElementById("s-ativos").textContent=r.ativos;
  document.getElementById("s-alocado").textContent=fmtBRL(r.totalAloc);
  document.getElementById("s-cob").textContent=r.cob.toFixed(1)+"%";
  document.getElementById("s-alert").textContent=r.alerts.length;
  document.getElementById("btn-pdf").disabled=false;
  document.getElementById("st-cob").className="stat "+(r.cob>99.99?"ok":"warn");
  var hasErr=false; for(var e=0;e<r.alerts.length;e++){ if(r.alerts[e].lvl==="err") hasErr=true; }
  document.getElementById("st-alert").className="stat "+(hasErr?"err":(r.alerts.length?"warn":"ok"));

  var h="<tr><th>Imobilizado</th><th>Chave</th><th>Classe</th><th class='num'>Valor consumido</th><th>Inventario</th><th>Data</th><th>GEAR</th></tr>";
  for(var i=0;i<r.out.length;i++){ var o=r.out[i];
    h+="<tr><td>"+esc(o.IMOBILIZADO)+"</td><td>"+esc(o.CHAVE)+"</td><td>"+esc(o.DESCRICAO_CLASSE_CUSTO||"")+"</td>"+
       "<td class='num'>"+(o.VALOR_CONSUMIDO?fmtBRL(o.VALOR_CONSUMIDO):"-")+"</td>"+
       "<td>"+esc(fmtCell(o.INVENTARIO))+"</td><td>"+esc(fmtDate(o.DATA))+"</td><td>"+esc(fmtCell(o.GEAR))+"</td></tr>";
  }
  document.getElementById("t-prev").innerHTML=h;

  h="<tr><th>Ativo</th><th>Classe</th><th class='num'>Demanda</th><th class='num'>Alocado</th><th class='num'>Diferenca</th><th>Status</th></tr>";
  for(var a=0;a<r.reconAC.length;a++){ var ac=r.reconAC[a]; var dif=ac.demanda-ac.alocado; var ok=Math.abs(dif)<EPS;
    h+="<tr><td>"+esc(ac.ativo)+"</td><td>"+esc(ac.classe)+"</td><td class='num'>"+fmtBRL(ac.demanda)+"</td>"+
       "<td class='num'>"+fmtBRL(ac.alocado)+"</td><td class='num'>"+fmtBRL(dif)+"</td>"+
       "<td><span class='badge "+(ok?"b-ok":"b-err")+"'>"+(ok?"OK":"FALTA")+"</span></td></tr>";
  }
  document.getElementById("t-ac").innerHTML=h;

  h="<tr><th>Chave</th><th class='num'>Saldo</th><th class='num'>Consumido</th><th class='num'>Restante</th><th>Status</th></tr>";
  for(var ch in r.chaveMeta){ var mm=r.chaveMeta[ch]; var cons=r.reconChave[ch]||0; var rem=mm.saldo-cons;
    var over=cons>mm.saldo+EPS; var neg=mm.saldo<0;
    h+="<tr><td>"+esc(ch)+"</td><td class='num "+(neg?"neg":"")+"'>"+fmtBRL(mm.saldo)+"</td><td class='num'>"+(cons?fmtBRL(cons):"-")+"</td>"+
       "<td class='num'>"+fmtBRL(rem)+"</td><td><span class='badge "+(over?"b-err":(neg?"b-warn":"b-ok"))+"'>"+(over?"ESTOURO":(neg?"NEGATIVO":"OK"))+"</span></td></tr>";
  }
  document.getElementById("t-ch").innerHTML=h;

  var al=document.getElementById("al-list");
  if(!r.alerts.length){ al.innerHTML="<div class='alert ok'>Nenhum alerta. Todas as demandas foram atendidas dentro dos saldos.</div>"; }
  else{ var ah=""; for(var x=0;x<r.alerts.length;x++){ ah+="<div class='alert "+(r.alerts[x].lvl==="err"?"err":"warn")+"'>"+esc(r.alerts[x].msg)+"</div>"; } al.innerHTML=ah; }
}

*/
function fmtCell(value){ return value==null||value===""?"-":value; }
function fmtDate(value){ if(value==null||value==="") return "-"; if(value instanceof Date) return value.toLocaleDateString("pt-BR"); return value; }
function normalizeHeader(value){
  return String(value==null?"":value).normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[^a-z0-9]/gi,"").toUpperCase();
}
function findValue(row,names){
  var keys=Object.keys(row.__n||{});
  for(var i=0;i<names.length;i++){
    var wanted=normalizeHeader(names[i]);
    for(var j=0;j<keys.length;j++){ if(keys[j]===wanted||keys[j].indexOf(wanted)>-1) return row.__n[keys[j]]; }
  }
  return "";
}
function headerScore(row){
  var names=[];
  for(var i=0;i<row.length;i++) names.push(normalizeHeader(row[i]));
  var expected=["PLANTAGLOBAL","LOCALIDADE","UAR","DESCRICAODAUAR","UNIDADEDEMEDIDA","NDEINVENTARIO","QUANTIDADE","FABRICANTE","NUMERODESERIE","MARCA","MODELO","DTDEENTRADAEMOPERACAO"];
  var score=0;
  for(var e=0;e<expected.length;e++){ for(var n=0;n<names.length;n++){ if(names[n]===expected[e]||names[n].indexOf(expected[e])>-1){ score++; break; } } }
  return score;
}
function findAnexoSheet(wb){
  var preferred=["RELATORIODEATIVOSINSTCOMOS","RELATORIODEATIVOSINSTSEMOS"];
  for(var p=0;p<preferred.length;p++){
    for(var ps=0;ps<wb.SheetNames.length;ps++){
      // IGNORA ABAS OCULTAS
      var hidden =
      wb.Workbook &&
      wb.Workbook.Sheets &&
      wb.Workbook.Sheets[ps]
      ? wb.Workbook.Sheets[ps].Hidden
      : 0;

      if(hidden !== 0){
      continue;
      }

      var preferredName=normalizeHeader(wb.SheetNames[ps]);
      if(preferredName.indexOf(preferred[p])===-1) continue;
      var preferredWs=wb.Sheets[wb.SheetNames[ps]], preferredRows=XLSX.utils.sheet_to_json(preferredWs,{header:1,raw:true,defval:null});
      var preferredBest=null, preferredLimit=Math.min(preferredRows.length,30);
      for(var pr=0;pr<preferredLimit;pr++){
        var preferredScore=headerScore(preferredRows[pr]||[]);
        if(!preferredBest||preferredScore>preferredBest.score) preferredBest={ws:preferredWs,row:pr,score:preferredScore,name:wb.SheetNames[ps]};
      }
      if(preferredBest&&preferredBest.score>=2) return preferredBest;
    }
  }
  var best=null;
  for(var s=0;s<wb.SheetNames.length;s++){
    // IGNORA ABAS OCULTAS
    var hidden =
    wb.Workbook &&
    wb.Workbook.Sheets &&
    wb.Workbook.Sheets[s]
    ? wb.Workbook.Sheets[s].Hidden
    : 0;

    if(hidden !== 0){
    continue;
    }

    var ws=wb.Sheets[wb.SheetNames[s]], arr=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
    var limit=Math.min(arr.length,20);
    for(var r=0;r<limit;r++){
      var score=headerScore(arr[r]||[]);
      if(!best||score>best.score) best={ws:ws,row:r,score:score,name:wb.SheetNames[s]};
    }
  }
  return best&&best.score>=2?best:null;
}
function readAnexo(wb){
  var found=findAnexoSheet(wb);
  if(!found) return [];
  var data=sheetRows(found.ws,found.row), out=[];
  for(var i=0;i<data.rows.length;i++){
    var row=data.rows[i];
    var values={
      planta:findValue(row,["PLANTA GLOBAL","PLANTA"]),
      localidade:findValue(row,["LOCALIDADE"]),
      uar:findValue(row,["UAR"]),
      descricaoUar:findValue(row,["DESCRICAO DA UAR","DESCRICAO UAR"]),
      unidade:findValue(row,["UNIDADE DE MEDIDA","UNIDADE"]),
      inventario:findValue(row,["N DE INVENTARIO","NUMERO DE INVENTARIO","INVENTARIO"]),
      quantidade:findValue(row,["QUANTIDADE","QTD"]),
      fabricante:findValue(row,["FABRICANTE"]),
      serie:findValue(row,["NUMERO DE SERIE","N DE SERIE","SERIE"]),
      marca:findValue(row,["MARCA"]),
      modelo:findValue(row,["MODELO"]),
      data:findValue(row,["DT DE ENTRADA EM OPERACAO","DATA DE ENTRADA EM OPERACAO","ENTRADA EM OPERACAO"])
    };
    var hasData=false; for(var key in values){ if(values[key]!==null&&values[key]!==""){ hasData=true; break; } }
    if(hasData) out.push(values);
  }
  return out;
}
function processar(){
  if(!state.anexo) throw new Error("Anexe uma planilha do Anexo 2 ou do Anexo 3 antes de continuar.");
  var out=readAnexo(state.anexo);
  if(!out.length) throw new Error("Nao encontrei dados do Anexo 2/3. Verifique se a planilha possui os cabecalhos da relacao de ativos e linhas preenchidas abaixo deles.");
  state.result={out:out}; render();
}
function render(){
  var r=state.result, table="<tr><th>Planta Global</th><th>Localidade</th><th>UAR</th><th>Descrição da UAR</th><th>Unidade de Medida</th><th>Nº de inventário</th><th>Quantidade</th><th>Fabricante</th><th>Número de Série</th><th>Marca</th><th>Modelo</th><th>Dt. entrada em operação</th></tr>";
  for(var i=0;i<r.out.length;i++){ var o=r.out[i]; table+="<tr><td>"+esc(o.planta)+"</td><td>"+esc(o.localidade)+"</td><td>"+esc(o.uar).substring(0,7)+"</td><td>"+esc(o.descricaoUar)+"</td><td>"+esc(o.unidade)+"</td><td>"+esc(o.inventario)+"</td><td>"+esc(o.quantidade)+"</td><td>"+esc(o.fabricante)+"</td><td>"+esc(o.serie)+"</td><td>"+esc(o.marca)+"</td><td>"+esc(o.modelo)+"</td><td>"+esc(fmtDate(o.data))+"</td></tr>"; }
  document.getElementById("t-prev").innerHTML=table;
  document.getElementById("results").classList.remove("hidden");
  document.getElementById("btn-pdf").disabled=false;
  document.getElementById("results").scrollIntoView({behavior:"smooth"});
}
function teoValue(id){ return (document.getElementById(id).value||"").trim(); }
function teoDate(id){
  var value=teoValue(id); if(!value) return "";
  var p=value.split("-"); return p.length===3?p[2]+"/"+p[1]+"/"+p[0]:value;
}
function teoRadio(name){ var el=document.querySelector('input[name="'+name+'"]:checked'); return el?el.value:""; }
function collectTeo(){
  return {
    pep:teoValue("teo-pep"), contrato:teoValue("teo-contrato"), descricao:teoValue("teo-descricao"),
    administrador:teoValue("teo-administrador"), empresa:teoValue("teo-empresa"), escopo:teoValue("teo-escopo"),
    endereco:teoValue("teo-endereco"), municipio:teoValue("teo-municipio"), bairro:teoValue("teo-bairro"),
    dataOperacao:document.getElementById("teo-varias-datas").checked?"Várias datas - Vide anexo I":teoDate("teo-data-operacao"),
    execucao:teoRadio("teo-execucao"), desativacao:teoRadio("teo-desativacao"),
    sabespNome:teoValue("teo-sabesp-nome"), sabespMatricula:teoValue("teo-sabesp-matricula"),
    contratadaNome:teoValue("teo-contratada-nome"), inspecao:teoDate("teo-inspecao")
  };
}
function addPdfField(doc,label,value,x,y,w){
  doc.setFont("helvetica","normal"); doc.text(label,x,y);
  var texto = doc.splitTextToSize(String(value||""),140);
  doc.setFont("helvetica","normal"); doc.text(texto,x+w,y);
}
function addPdfSection(doc,title,y){
  doc.setFillColor(218,218,218); doc.rect(12,y,186,7,"F");
  doc.setFont("helvetica","bold"); doc.setFontSize(9); doc.text(title,14,y+4.7); return y+7;
}
function drawPdfFooter(doc,landscape){
  var page=doc.internal.getNumberOfPages();
  doc.setFontSize(7); doc.setFont("helvetica","normal");
  if(!landscape){
    var x=12, y=265, w=186, h=22;
    doc.setDrawColor(30,30,30); doc.setLineWidth(.25); doc.rect(x,y,w,h);
    doc.line(x+48,y,x+48,y+10); doc.line(x,y+10,x+w,y+10);
    doc.setFontSize(7); doc.text("Código do Formulário",x+2,y+5); doc.text("Nome do Formulário",x+50,y+5);
    doc.setFontSize(8); doc.text("FE-ATIVOS0005 - V.2",x+2,y+9); doc.text("TEO - TERMO DE ENTRADA EM OPERAÇÃO",x+50,y+9);
    doc.setFontSize(7); doc.text("Vinculado ao Instrumento",x+2,y+15); doc.setFontSize(7); doc.text("DOCUMENTAÇÃO TÉCNICA PARA UNITIZAÇÃO E DESATIVAÇÃO DE ATIVOS - PE-ATIVOS0003.",x+2,y+20);
  } else {
    var ly=202; doc.text("Código do Formulário: FE-ATIVOS0005 - V.2",14,ly); doc.text("TEO - TERMO DE ENTRADA EM OPERAÇÃO",135,ly); doc.text("Vinculado ao Instrumento: DOCUMENTAÇÃO TÉCNICA PARA UNITIZAÇÃO E DESATIVAÇÃO DE ATIVOS - PE-ATIVOS0003.",14,ly+5); doc.text("Página "+page,265,ly+10);
  }
}
function removeYellowBackground(dataUrl){
  return new Promise(function(resolve,reject){
    var image=new Image();
    image.onload=function(){
      var canvas=document.createElement("canvas"), context;
      canvas.width=image.naturalWidth; canvas.height=image.naturalHeight; context=canvas.getContext("2d");
      context.drawImage(image,0,0);
      var pixels=context.getImageData(0,0,canvas.width,canvas.height), data=pixels.data;
      for(var i=0;i<data.length;i+=4){
        var red=data[i], green=data[i+1], blue=data[i+2];
        if(red>170&&green>145&&blue<125&&red>blue*1.5&&green>blue*1.25) data[i+3]=0;
      }
      context.putImageData(pixels,0,0); resolve(canvas.toDataURL("image/png"));
    };
    image.onerror=reject; image.src=dataUrl;
  });
}

const SABESP_LOGO =
'data:image/bmp;base64,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';

function drawSabespPdfLogo(doc){
  if(!state.sabespLogoData) return;
  try{ doc.addImage(state.sabespLogoData,"PNG",177,10,20,25); }
  catch(error){ console.warn("Logo SABESP nao inserida",error); }
}
function loadSabespLogo(){
    state.sabespLogoData = SABESP_LOGO;

    var headerImage =
        document.getElementById("sabesp-header-image");

    if(headerImage){
        headerImage.src = SABESP_LOGO;
    }

    return Promise.resolve(true);
}
function addAnexoHeader(doc){
  doc.setDrawColor(30,30,30); doc.setLineWidth(.25); doc.rect(7,5,283,18);
  // divisão do logo
  doc.line(27,5,27,23);
  // divisão do número do anexo
  doc.line(263,5,263,23);
  // linha horizontal
  doc.line(27,14,263,14);
  doc.setFont("helvetica","normal");
  doc.text("Nome do Anexo",30,9);
  doc.setFont("helvetica","bold");
  doc.text("Relação de Ativos - Entrada em Operação",30,13);
  doc.setFont("helvetica","normal");
  doc.text("Vinculado ao Instrumento",30,17);
  doc.setFont("helvetica","bold");
  doc.text("FE-ATIVOS0005 V.1 - TEO - Termo de Entrada em Operação",30,21);
  doc.line(263,5,263,23); doc.setFontSize(7); doc.text("Número do Anexo",265,10); doc.setFontSize(9); doc.text("0001",276,17,{align:"center"});
  doc.setFillColor(14,174,224); doc.roundedRect(7,25,283,15,2,2,"F"); doc.setTextColor(255,255,255); doc.setFontSize(15); doc.text("Anexo 1 - Relação de Ativos - Entrada em Operação",16,35);
}
function gerarPdf(){
  if(!state.result) return;
  var d=collectTeo();
  var doc=new window.jspdf.jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
  doc.setTextColor(20,20,20); doc.setFont("helvetica","normal"); doc.setFontSize(9);
  doc.setFont("helvetica","bold"); doc.setFontSize(16); doc.text("TEO - Termo de Entrada em Operação",105,14,{align:"center"});
  if(state.logoData){ try{ var alturaMax = 12;

var larguraLogo =
    (state.logoWidth / state.logoHeight)
    * alturaMax;

doc.addImage(
    state.logoData,
    state.logoType || "PNG",
    12, 
    17,
    larguraLogo,
    alturaMax
); }catch(e){ console.warn("Logo nao inserida",e); } }
  drawSabespPdfLogo(doc);
  var y=35; y=addPdfSection(doc,"1. Identificação da Obra",y); doc.setFontSize(8);
  var rows=[[["PEP:",d.pep],["Nº. Contrato:",d.contrato]],[["Descrição do PEP:",d.descricao]],[["Administrador do Contrato:",d.administrador]],[["Empresa Contratada:",d.empresa]],[["Escopo do Contrato:",d.escopo]],[["Endereço:",d.endereco]],[["Município:",d.municipio],["Bairro:",d.bairro]],[["Data Entrada em Operação:",d.dataOperacao],["Execução:",d.execucao]],[["Houve desativação?",d.desativacao]]];
  for(var i=0;i<rows.length;i++){
    var h=rows[i].length>1?7:7; 
    // Escopo do contrato
    if(rows[i][0][0] === "Escopo do Contrato:"){
    var linhas = doc.splitTextToSize(
    rows[i][0][1] || "",140);
  h = Math.max(10,linhas.length * 4 + 4);
    }
    
    doc.setDrawColor(190,190,190);
    doc.rect(12,y,186,h);
    doc.setDrawColor(190,190,190); doc.rect(12,y,186,h);
    if(rows[i].length===1){ addPdfField(doc,rows[i][0][0],rows[i][0][1],14,y+5,42); }
    else { doc.line(108,y,108,y+h); addPdfField(doc,rows[i][0][0],rows[i][0][1],14,y+5,42); addPdfField(doc,rows[i][1][0],rows[i][1][1],110,y+5,28); }
    y+=h;
  }
  y+=7; y=addPdfSection(doc,"2. Responsável Sabesp",y); doc.setDrawColor(190,190,190); doc.rect(12,y,186,8); addPdfField(doc,"Nome:",d.sabespNome,14,y+5,42); y+=8; doc.rect(12,y,186,8); addPdfField(doc,"Matrícula:",d.sabespMatricula,14,y+5,42);
  y+=14; y=addPdfSection(doc,"3. Responsável Contratada",y); doc.rect(12,y,186,8); addPdfField(doc,"Nome:",d.contratadaNome,14,y+5,42); y+=8; doc.rect(12,y,186,8); addPdfField(doc,"Data da Inspeção:",d.inspecao,14,y+5,42);
  y+=14; doc.setFont("helvetica","normal"); doc.setFontSize(8.5); var text="Os responsáveis abaixo declaram que os ativos relacionados no Anexo I, instalados na área de concessão da Companhia de Saneamento Básico do Estado de São Paulo S.A. (SABESP) e objeto do referido Contrato, foram devidamente implementados, comissionados, conectados ao Sistema da Concessionária."; doc.text(doc.splitTextToSize(text,186),12,y); y+=15; doc.text(doc.splitTextToSize("Ressalta-se, ainda, que todos os custos diretos associados aos ativos mencionados no Anexo I encontram-se devidamente apropriados no PEP.",186),12,y); y=238;
  doc.line(12,y,87,y); doc.line(112,y,187,y); doc.text("Responsável Contratada:",14,y+8); doc.text("Responsável Sabesp:",114,y+8); doc.text("Data: ____/____/_______",14,y+14); doc.text("Data: ____/____/_______",114,y+14); drawPdfFooter(doc);
  doc.addPage("a4","landscape"); addAnexoHeader(doc); if(state.sabespLogoData){
doc.addImage(
state.sabespLogoData,
"PNG",
10,
5,
14,
18
);
} 
doc.setFontSize(8); doc.setFont("helvetica","normal");
  var body=[]; for(var j=0;j<state.result.out.length;j++){ var o=state.result.out[j]; body.push([String(o.planta||""),String(o.localidade||""),String(o.uar||"").substring(0,7),String(o.descricaoUar||""),String(o.unidade||""),String(o.inventario||""),String(o.quantidade||""),String(o.fabricante||""),String(o.serie||""),String(o.marca||""),String(o.modelo||""),String(fmtDate(o.data))]); }
  doc.autoTable({startY:45,margin:{left:7,right:7},head:[["Planta Global","Localidade","UAR","Descrição da UAR","Unidade de Medida","Nº de inventário","Quantidade","Fabricante (se aplicável)","Número de Série (se aplicável)","Marca (se aplicável)","Modelo (se aplicável)","Dt. de entrada em operação"]],body:body,theme:"grid",styles:{fontSize:6.5,fontStyle:"normal",textColor:[0,0,0],cellPadding:0.8,overflow:"linebreak",valign:"middle"},headStyles:{fillColor:[0,58,89],textColor:[255,255,255],fontStyle:"bold",fontSize:6.5},columnStyles:{
    0:{cellWidth:15}, // Planta Global
    1:{cellWidth:15}, // Localidade
    2:{cellWidth:17}, // UAR
    3:{cellWidth:80}, // Descrição UAR
    4:{cellWidth:15}, // Unidade
    5:{cellWidth:20}, // Inventário
    6:{cellWidth:16}, // Quantidade
    7:{cellWidth:20}, // Fabricante
    8:{cellWidth:25}, // Série
    9:{cellWidth:20}, // Marca
    10:{cellWidth:20}, // Modelo
    11:{cellWidth:20} // Data Entrada
}}); drawPdfFooter(doc,true);
  var base=(d.pep||"TEO").replace(/[^a-z0-9_-]+/gi,"_"); doc.save(base+"_TEO.pdf");
}

function showMsg(text){ var message=document.getElementById("msg"); message.textContent=text; message.className="msg show error"; }
function clearMsg(){ document.getElementById("msg").className="msg"; }  
function checkReady(){ document.getElementById("btn-run").disabled=!state.anexo; }
function loadFile(file){
  clearMsg();
  readWorkbook(file).then(function(wb){
    state.anexo=wb;
    var drop=document.getElementById("drop-anexo"), meta=document.getElementById("meta-anexo");
    var found=findAnexoSheet(wb);
    drop.classList.add("loaded"); meta.classList.remove("hidden");
    meta.textContent=found?"OK: "+file.name+" - aba: "+found.name+" - cabecalho: linha "+(found.row+1):"Arquivo carregado, mas nao localizei o cabecalho do Anexo 2/3";
    checkReady();
  }).catch(function(err){
    var details=String(err&&err.message||err);
    if(/Encrypted file|EncryptionInfo|ECMA-376/i.test(details)){
      showMsg("Este arquivo esta protegido por criptografia ou rotulo de confidencialidade do Office. Abra-o no Excel e salve uma copia sem protecao/criptografia antes de anexar.");
    } else showMsg("Erro ao ler "+file.name+": "+details);
  });
}
function bindWorkbookDrop(){
  var drop=document.getElementById("drop-anexo"), input=document.getElementById("file-anexo");
  drop.addEventListener("click",function(){ input.click(); });
  drop.addEventListener("dragover",function(event){ event.preventDefault(); drop.classList.add("drag"); });
  drop.addEventListener("dragleave",function(){ drop.classList.remove("drag"); });
  drop.addEventListener("drop",function(event){ event.preventDefault(); drop.classList.remove("drag"); if(event.dataTransfer.files[0]) loadFile(event.dataTransfer.files[0]); });
  input.addEventListener("change",function(event){ if(event.target.files[0]) loadFile(event.target.files[0]); });
}

/*
// ---------- estilo e download antigos do LOG ----------
var LARANJA={fill:{patternType:"solid",fgColor:{rgb:"FFF26C3F"}},font:{bold:true,color:{rgb:"FFFFFFFF"}}};
// Re-aplica o laranja na linha de cabecalho (a lib perde o fill no round-trip por diferenca de formato).
function carimbaCabecalho(ws){
  if(!ws||!ws["!ref"]) return;
  var range=XLSX.utils.decode_range(ws["!ref"]);
  for(var c=range.s.c;c<=range.e.c;c++){
    var addr=XLSX.utils.encode_cell({r:range.s.r,c:c});
    if(ws[addr]) ws[addr].s=LARANJA;
  }
}

// ---------- download antigo ----------
function baixar(){
  var r=state.result; if(!r) return;
  var wb=state.log;  // mesmo workbook que o usuario subiu

  // 1) preenche a guia Modelo Upload IN-PLACE (mantem Chaves com Saldo e a formatacao intactas)
  var wsModelo=wb.Sheets[r.modeloName];
  fillModelo(wsModelo, r.alloc, r.assetMeta);

  // re-carimba o cabecalho laranja das guias de dados (preserva o visual do LOG original)
  for(var si=0;si<wb.SheetNames.length;si++){ carimbaCabecalho(wb.Sheets[wb.SheetNames[si]]); }

  // 2) (re)cria a guia Conferencia
  var confName="Conferencia";
  if(wb.SheetNames.indexOf(confName)>-1){
    delete wb.Sheets[confName];
    wb.SheetNames.splice(wb.SheetNames.indexOf(confName),1);
  }
  var conf=[["CONFERENCIA ATIVO x CLASSE"],["Ativo","Classe","Demanda","Alocado","Diferenca"]];
  for(var a=0;a<r.reconAC.length;a++){ var ac=r.reconAC[a]; conf.push([ac.ativo,ac.classe,ac.demanda,ac.alocado,ac.demanda-ac.alocado]); }
  conf.push([]); conf.push(["CONFERENCIA POR CHAVE (saldo da coluna J)"]); conf.push(["Chave","Saldo disponivel","Consumido","Restante"]);
  for(var ch in r.chaveMeta){ var mm=r.chaveMeta[ch]; var c=r.reconChave[ch]||0; conf.push([ch,mm.saldo,c,mm.saldo-c]); }
  conf.push([]); conf.push(["ALERTAS"]);
  if(!r.alerts.length){ conf.push(["OK","Nenhum alerta. Todas as demandas atendidas dentro dos saldos."]); }
  for(var x=0;x<r.alerts.length;x++){ conf.push([r.alerts[x].lvl.toUpperCase(),r.alerts[x].msg]); }
  var ws2=XLSX.utils.aoa_to_sheet(conf);
  ws2["!cols"]=[{wch:50},{wch:18},{wch:16},{wch:16},{wch:14}];
  // cabecalhos/titulos da Conferencia em PRETO com fonte branca
  var preto={fill:{patternType:"solid",fgColor:{rgb:"FF000000"}},font:{bold:true,color:{rgb:"FFFFFFFF"}}};
  var titleRows=[]; for(var ti=0;ti<conf.length;ti++){ var first=conf[ti][0]; if(first==="Ativo"||first==="Chave"||(typeof first==="string"&&first.indexOf("CONFERENCIA")===0)||first==="ALERTAS"){ titleRows.push(ti); } }
  for(var tr=0;tr<titleRows.length;tr++){
    var rr=titleRows[tr];
    for(var cc=0;cc<5;cc++){
      var addr=XLSX.utils.encode_cell({r:rr,c:cc});
      if(ws2[addr]) ws2[addr].s=preto;
    }
  }
  XLSX.utils.book_append_sheet(wb,ws2,confName);

  // 3) nome de saida derivado do arquivo original
  var base=(state.logName||"consulta_log_unitizacao.xlsx").replace(/\.(xlsx|xls)$/i,"");
  XLSX.writeFile(wb, base+"_PREENCHIDO.xlsx", {cellStyles:true});
}

*/
/* ---------- UI ---------- */
// impede o navegador de abrir/baixar o arquivo se ele for solto fora da caixa
window.addEventListener("dragover",function(e){ e.preventDefault(); }, false);
window.addEventListener("drop",function(e){ e.preventDefault(); }, false);

/* ---------- toggle de tema claro/escuro ---------- */



document.addEventListener("DOMContentLoaded",function(){
  // estado inicial dos botoes do toggle
  try{ saved=localStorage.getItem("logsync-theme")||"dark"; }catch(e){}
  loadSabespLogo();
  


  bindWorkbookDrop();
  document.getElementById("teo-varias-datas").addEventListener("change",function(){
    if(this.checked) document.getElementById("teo-data-operacao").value="";
  });
  document.getElementById("teo-logo").addEventListener("change",function(e){
    var file=e.target.files[0]; if(!file) return;
    var reader=new FileReader(); reader.onload=function(ev){
      var img = new Image();

img.onload = function(){

    state.logoWidth  = img.width;
    state.logoHeight = img.height;

};

img.src = ev.target.result;

      removeYellowBackground(ev.target.result).then(function(cleanData){ state.logoData=cleanData; state.logoType="PNG"; document.getElementById("teo-logo-name").textContent=file.name; }).catch(function(){ state.logoData=ev.target.result; state.logoType=file.type.indexOf("png")>-1?"PNG":"JPEG"; document.getElementById("teo-logo-name").textContent=file.name; });
    };
    reader.readAsDataURL(file);
  });
  document.getElementById("btn-run").addEventListener("click",function(){
    clearMsg();
    try{ processar(); document.getElementById("results").scrollIntoView({behavior:"smooth"}); }
    catch(err){ showMsg("Erro ao processar: "+err.message); console.error(err); }
  });
  document.getElementById("btn-pdf").addEventListener("click",function(){
    loadSabespLogo().then(function(){ try{ gerarPdf(); }catch(err){ showMsg("Erro ao gerar o PDF: "+err.message); console.error(err); } }).catch(function(err){ showMsg("Erro ao carregar o logo: "+err.message); });
  });
  document.getElementById("btn-reset").addEventListener("click",function(){ location.reload(); });
});
</script>
</body>
</html>
