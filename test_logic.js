// Teste: replica as fórmulas do Excel e valida contra os valores esperados
// F39: =IF(OR(C39="",D39="",E39=""),"",C39*D39*E39)
// D46: =IF(COUNT(F39:F44)=0,"",IF(MAX(F39:F44)>=33,"Alto",IF(MAX(F39:F44)>=21,"Médio","Baixo")))
// Nº:  Trim(Config!C3) & "-" & Config!C4 & "-" & Format(num,"0000")

function nivelRisco(s,e,c){ if(s==null||e==null||c==null||s===''||e===''||c==='') return null; return Number(s)*Number(e)*Number(c); }
function riscoGlobal(vals){ const n=vals.filter(v=>v!==null&&v!==''); if(!n.length) return null; const mx=Math.max(...n); return mx>=33?'Alto':mx>=21?'Médio':'Baixo'; }
function formatNum(p,a,n){ return `${String(p).trim()}-${a}-${String(n).padStart(4,'0')}`; }

let fails = 0;
function check(name, got, want){
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`${ok?'✅':'❌'} ${name}: got=${JSON.stringify(got)} want=${JSON.stringify(want)}`);
  if(!ok) fails++;
}

// Cenário 1 — célula incompleta (fórmula devolve vazio)
check('F-célula incompleta', nivelRisco(3,2,''), null);

// Cenário 2 — risco baixo: 2×2×2=8 → verde
check('F-cálculo 2×2×2', nivelRisco(2,2,2), 8);
check('D46 baixo', riscoGlobal([8]), 'Baixo');

// Cenário 3 — moderado: 3×4×2=24 → amarelo (>=21)
check('D46 médio (24)', riscoGlobal([24]), 'Médio');
check('D46 médio (21 limite)', riscoGlobal([21]), 'Médio');

// Cenário 4 — alto: 4×4×3=48 → vermelho (>=33)
check('D46 alto (48)', riscoGlobal([48]), 'Alto');
check('D46 alto (33 limite)', riscoGlobal([33]), 'Alto');

// Cenário 5 — máximo de várias linhas: [8,24,48] → Alto (MAX)
check('D46 usa MAX', riscoGlobal([8,24,48]), 'Alto');
check('D46 MAX com baixo+alto', riscoGlobal([5,36]), 'Alto');

// Cenário 6 — sem riscos avaliados → vazio
check('D46 vazio', riscoGlobal([null,'',null]), null);

// Cenário 7 — numeração: Config!C3=' AT' (com espaço), C4=2026, C5=1 → AT-2026-0001
check('Numeração 1', formatNum(' AT','2026',1), 'AT-2026-0001');
check('Numeração 42', formatNum('AT','2026',42), 'AT-2026-0042');
check('Numeração 1234', formatNum('AT','2026',1234), 'AT-2026-1234');

process.exit(fails ? 1 : 0);
