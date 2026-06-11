import { useState, useRef, useEffect } from "react";
import * as mammoth from "mammoth";
import Papa from "papaparse";
import PptxGenJS from "pptxgenjs";

const C = {
  bg:'#0B1120', surface:'#111827', card:'#1F2937', border:'#374151',
  accent:'#F97316', text:'#F9FAFB', muted:'#9CA3AF', faint:'#6B7280',
  success:'#10B981', warn:'#F59E0B', error:'#EF4444', blue:'#3B82F6',
};
const st = {
  app:{ display:'flex', height:'100vh', background:C.bg, color:C.text, fontFamily:"'Inter',sans-serif", overflow:'hidden' },
  sidebar:{ width:220, background:C.surface, borderRight:'1px solid #1F2937', display:'flex', flexDirection:'column', flexShrink:0 },
  main:{ flex:1, overflow:'auto', display:'flex', flexDirection:'column' },
  header:{ padding:'18px 28px 14px', borderBottom:'1px solid #1F2937', background:C.surface, flexShrink:0 },
  content:{ padding:28, flex:1 },
  card:{ background:C.card, borderRadius:12, border:`1px solid ${C.border}`, padding:20 },
  input:{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:8, padding:'10px 14px', color:C.text, fontSize:14, width:'100%', outline:'none', boxSizing:'border-box' },
  textarea:{ background:C.surface, border:`1px solid ${C.border}`, borderRadius:8, padding:'10px 14px', color:C.text, fontSize:14, width:'100%', outline:'none', boxSizing:'border-box', resize:'vertical' },
  label:{ fontSize:11, fontWeight:700, color:C.muted, marginBottom:6, display:'block', textTransform:'uppercase', letterSpacing:'0.6px' },
};
const btn = (v='primary') => ({
  display:'inline-flex', alignItems:'center', justifyContent:'center', gap:7,
  padding: v==='sm'?'5px 13px':'10px 20px', borderRadius:8,
  fontSize: v==='sm'?12:14, fontWeight:600, cursor:'pointer',
  border: v==='outline'?`1px solid ${C.border}`:'none',
  background: v==='primary'?C.accent: v==='ghost'?'#374151': v==='outline'?'transparent':'#374151',
  color: v==='outline'?C.muted:'#fff',
});
const badge = (t) => ({
  display:'inline-flex', alignItems:'center', padding:'2px 9px', borderRadius:20, fontSize:11, fontWeight:600,
  background: t==='prospect'?'#1D4ED815':'#05966915',
  color: t==='prospect'?'#60A5FA':'#34D399',
  border:`1px solid ${t==='prospect'?'#1D4ED840':'#05966940'}`,
});
const navItem = (a) => ({
  display:'flex', alignItems:'center', gap:9, padding:'9px 12px', borderRadius:8,
  cursor:'pointer', fontSize:13, fontWeight:a?600:400,
  color:a?C.accent:C.muted, background:a?'#374151':'transparent', marginBottom:2,
});
const slideIcons = { 'cover':'🎯','executive-summary':'📋','current-state':'📊','gap-analysis':'🔍','solution':'💡','dsp-strategy':'📡','ctv-strategy':'📺','timeline':'🗓','investment':'💰','next-steps':'🚀','month-in-review':'📅','kpi-performance':'📈','whats-working':'✅','needs-attention':'⚠️','strategy':'🎯' };

const REPORT_TYPES = [
  { id:'sb_search_term', label:'SB Search Term Report', cat:'Sponsored Brands' },
  { id:'sb_kw_placement', label:'SB Keyword Placement', cat:'Sponsored Brands' },
  { id:'sp_placement', label:'SP Placement Report', cat:'Sponsored Products' },
  { id:'sp_advertised', label:'SP Advertised Product', cat:'Sponsored Products' },
  { id:'sp_search_term', label:'SP Search Term Report', cat:'Sponsored Products' },
  { id:'beta_product', label:'Beta Report by Product', cat:'DSP / Beta' },
  { id:'beta_target', label:'Beta Report by Target', cat:'DSP / Beta' },
  { id:'beta_campaign', label:'Beta Report by Campaign', cat:'DSP / Beta' },
  { id:'sc_sales_asin', label:'SC Total Sales by ASIN', cat:'Seller Central' },
  { id:'ss_search', label:'SmartScout: Search Terms', cat:'SmartScout' },
  { id:'ss_market', label:'SmartScout: Market Share', cat:'SmartScout' },
  { id:'ss_products', label:'SmartScout: Products', cat:'SmartScout' },
  { id:'ss_brand', label:'SmartScout: Brand Monthly', cat:'SmartScout' },
];

const CAT_COLORS = { 'Sponsored Brands':'#3B82F6','Sponsored Products':'#8B5CF6','DSP / Beta':'#F97316','Seller Central':'#10B981','SmartScout':'#F59E0B' };

const DEMO_CLIENTS = [
  { id:'d1', name:'NovaBrand Nutrition', type:'prospect', contact:'Sarah Chen', email:'sarah@novabrand.com', categories:'Health & Wellness, Sports Nutrition', goals:'Break top 3 in protein powder, improve ROAS to 4x', sessions:[], createdAt:'2026-05-01' },
  { id:'d2', name:'PureHome Essentials', type:'active', contact:'Marcus Webb', email:'marcus@purehome.com', categories:'Home & Garden, Cleaning', goals:'Grow Amazon revenue 40% YoY, expand to 3 new subcategories', sessions:[{ id:'s1', type:'monthly', date:'2026-05-01', title:'May 2026 Performance Review', slideCount:7 }], createdAt:'2025-09-01' },
];

const FALLBACK_ANALYSIS = {
  painPoints:['Low visibility vs. category competitors','Poor ROAS on Sponsored Products (~2.1x)','No DSP or programmatic strategy'],
  goals:['Increase market share by 25%','Achieve 4x ROAS within 6 months','Launch DSP retargeting'],
  currentStrategy:['Basic Sponsored Products with manual bidding','No external traffic or DSP','Minimal branded search defense'],
  gaps:['No DSP, programmatic, or CTV/STV strategy','Missing mid-funnel and retargeting coverage','Competitors outspending on programmatic'],
  keyMetrics:['ROAS: ~2.1x','Monthly ad spend: ~$15K','Category rank: top 10, flat'],
  budgetTimeline:['Ready to scale to $30K+/month','Decision in next 30 days'],
  decisionFactors:['Proven results in category','Transparent reporting','No long-term contracts'],
  competitiveContext:['3 main competitors dominate DSP impressions'],
  quickWins:['Launch DSP retargeting — product page visitors not being recaptured','Restructure SP campaigns for mid-funnel coverage'],
  summary:'A growth-stage brand with real budget and urgency, stalled by a strategy that stopped evolving. They have the goals, the spend, and the desire to win — they need a partner with programmatic depth.',
};

const buildFallbackSlides = (n) => ([
  { id:'sl-1', type:'cover', title:n, headline:'Amazon Advertising Strategy Proposal', bullets:[], notes:'Set the stage.' },
  { id:'sl-2', type:'executive-summary', title:'Executive Summary', headline:'Strong brand. Under-leveraged channel. Clear path forward.', bullets:['Growth has plateaued despite meaningful ad investment','Competitors pulling ahead with programmatic strategies not yet activated','BTR Media can close the gap in 90 days'], notes:'Show you understand their business before selling anything.' },
  { id:'sl-3', type:'current-state', title:'Where You Are Today', headline:'The foundation is solid — but the ceiling is low without DSP.', bullets:['Sponsored Products running, limited to manual bidding','No programmatic reach to retarget, conquest, or expand','Share of voice eroding as competitors invest in DSP'], notes:'Mirror back what they told you.' },
  { id:'sl-4', type:'gap-analysis', title:'The Gaps Holding You Back', headline:'Three gaps are creating a performance ceiling you cannot break through alone.', bullets:['No DSP — product page visitors leave and never come back','No CTV/STV — competitors building brand equity off Amazon','Mid-funnel gap — brand-aware shoppers not being captured'], notes:'Be direct. Urgency comes from clarity.' },
  { id:'sl-5', type:'solution', title:'The BTR Media Solution', headline:'Full-funnel Amazon advertising — built around your goals.', bullets:['Amazon DSP: Retarget, conquest, and reach audiences at every stage','Sponsored Ads rebuilt for maximum mid-funnel and branded coverage','CTV/STV: Drive awareness with Amazon purchase-data audiences','Dedicated AM + DSP specialist, weekly reporting'], notes:'Connect every solution to a gap you named.' },
  { id:'sl-6', type:'dsp-strategy', title:'DSP Strategy', headline:'Recapture the demand you are generating but not converting.', bullets:['Retargeting: Re-engage product page visitors within 30 days','Conquest: Appear on competitor ASINs and category pages','Lifestyle audiences: Reach high-intent shoppers before they search','Lookalike scaling: Expand your best customer profiles'], notes:'DSP is your biggest differentiator.' },
  { id:'sl-7', type:'ctv-strategy', title:'CTV/STV Opportunity', headline:'Your competitors are building brand equity off Amazon. You should be too.', bullets:['Streaming TV ads served to Amazon verified purchase-data audiences','Drives branded search lift and top-of-funnel awareness','Tracked through Amazon Attribution','Target cord-cutters in your exact customer demographic'], notes:'Position as a competitive moat.' },
  { id:'sl-8', type:'timeline', title:'90-Day Launch Plan', headline:'From kickoff to full-funnel activation in 90 days.', bullets:['Days 1-14: Audit, SP restructure, DSP setup + audience build','Days 15-30: SP rebuild live + first DSP campaigns launched','Days 31-60: Optimize, layer conquest and retargeting tiers','Days 61-90: CTV/STV launch, full-funnel reporting dashboard live'], notes:'Concrete timelines remove fear of the unknown.' },
  { id:'sl-9', type:'investment', title:'Investment', headline:'Structured to grow with your results — not front-loaded.', bullets:['Management fee based on total ad spend under management','No long-term contracts','Weekly optimization + bi-weekly strategy calls','Dedicated AM + DSP specialist from day one'], notes:'Keep high-level unless specific pricing was requested.' },
  { id:'sl-10', type:'next-steps', title:'Next Steps', headline:'Ready to build this together?', bullets:['Share Amazon Seller/Vendor Central access for full audit','Send over SmartScout report for competitive analysis','Schedule kickoff call within 5 business days','BTR team ready to activate within 1 week of agreement'], notes:'End with clear actions. Make it simple to say yes.' },
]);

const buildFallbackMonthly = (n) => ([
  { id:'sl-1', type:'cover', title:`${n} — Monthly Review`, headline:new Date().toLocaleDateString('en-US',{month:'long',year:'numeric'}), bullets:[], notes:'' },
  { id:'sl-2', type:'month-in-review', title:'Month in Review', headline:'Key wins, optimizations, and what we moved on.', bullets:['[Top campaign win — replace with actuals]','[Budget efficiency improvement]','[New tests or activations launched]'], notes:'' },
  { id:'sl-3', type:'kpi-performance', title:'KPI Performance vs. Goals', headline:'Tracking toward target — here is where we stand.', bullets:['Revenue: $X vs. $Y goal','ROAS: X.Xx vs. Y.Yx target','New to Brand: X%','Share of Voice: X%'], notes:'Pull actuals before presenting.' },
  { id:'sl-4', type:'whats-working', title:"What's Working", headline:'Double down on these.', bullets:['[Top campaign / tactic]','[Best converting audience]','[Top ASIN or keyword]'], notes:'' },
  { id:'sl-5', type:'needs-attention', title:'Needs Attention', headline:'Active optimizations in progress.', bullets:['[Underperforming area — action taken]','[Gap vs. target — close plan]','[Competitive pressure — response]'], notes:'' },
  { id:'sl-6', type:'strategy', title:'Next Month Priorities', headline:'Three things we are focused on next month.', bullets:['Expand DSP retargeting to new audience segments','Test updated creative formats','Increase allocation to highest-ROAS campaigns'], notes:'' },
  { id:'sl-7', type:'next-steps', title:'Action Items', headline:'What happens between now and next month.', bullets:['BTR: Optimizations live by [date]','Client: Updated creative assets by [date]','Both: Bi-weekly check-in on [date]'], notes:'' },
]);

export default function BTRDeckStudio() {
  const [screen, setScreen] = useState('clients');
  const [clients, setClients] = useState(DEMO_CLIENTS);
  const [selId, setSelId] = useState(null);
  const [sessionType, setSessionType] = useState('discovery');
  const [step, setStep] = useState(0);
  const [transcript, setTranscript] = useState('');
  const [fileName, setFileName] = useState('');
  const [analyzing, setAnalyzing] = useState(false);
  const [analysis, setAnalysis] = useState(null);
  const [uploadedReports, setUploadedReports] = useState({});
  const [pendingRpt, setPendingRpt] = useState(null);
  const [analyzingReports, setAnalyzingReports] = useState(false);
  const [reportAnalysis, setReportAnalysis] = useState(null);
  const [chatHistory, setChatHistory] = useState([]);
  const [chatInput, setChatInput] = useState('');
  const [chatLoading, setChatLoading] = useState(false);
  const [generating, setGenerating] = useState(false);
  const [slides, setSlides] = useState([]);
  const [editId, setEditId] = useState(null);
  const [exporting, setExporting] = useState(false);
  const transcriptRef = useRef(null);
  const reportFileRef = useRef(null);
  const chatEndRef = useRef(null);

  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior:'smooth' }); }, [chatHistory]);

  const client = clients.find(c => c.id === selId);

  const handleFile = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    setFileName(f.name);
    if (f.name.match(/\.docx?$/i)) { const buf=await f.arrayBuffer(); const r=await mammoth.extractRawText({arrayBuffer:buf}); setTranscript(r.value); }
    else { setTranscript(await f.text()); }
  };

  const doAnalyze = async () => {
    setAnalyzing(true);
    try {
      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:1000,
          system:`You are a senior Amazon advertising strategist at BTR Media. Analyze discovery call transcripts for sales intelligence. Return ONLY valid JSON, no markdown.`,
          messages:[{ role:'user', content:`Analyze this transcript for ${client?.name}. Return JSON with: painPoints, goals, currentStrategy, gaps, keyMetrics, budgetTimeline, decisionFactors, competitiveContext, quickWins (each an array of strings), and summary (2-3 sentences).\n\nTRANSCRIPT:\n${transcript}` }]
        })
      });
      const d = await res.json();
      setAnalysis(JSON.parse((d.content?.[0]?.text||'{}').replace(/```json|```/g,'').trim()));
      setStep(1);
    } catch { setAnalysis(FALLBACK_ANALYSIS); setStep(1); }
    setAnalyzing(false);
  };

  const triggerReportUpload = (rptId) => { setPendingRpt(rptId); setTimeout(()=>reportFileRef.current?.click(),50); };

  const handleReportFile = async (e) => {
    const f = e.target.files[0]; if (!f || !pendingRpt) return;
    e.target.value = '';
    const rtype = REPORT_TYPES.find(r => r.id === pendingRpt);
    try {
      await new Promise((resolve) => {
        Papa.parse(f, {
          header:true, skipEmptyLines:true, dynamicTyping:true,
          complete:(results) => {
            const headers = results.meta.fields || [];
            const allRows = results.data;
            const rows = allRows.slice(0, 150);
            const dataText = [
              `=== ${rtype.label} (${rtype.cat}) ===`,
              `Total rows: ${allRows.length} | Showing first ${rows.length}`,
              `Columns: ${headers.join(', ')}`,
              '---',
              rows.map(row => headers.map(h=>`${h}: ${row[h]??''}`).join(' | ')).join('\n')
            ].join('\n');
            setUploadedReports(prev => ({ ...prev, [pendingRpt]:{ label:rtype.label, cat:rtype.cat, fileName:f.name, rowCount:allRows.length, headers, dataText } }));
            setPendingRpt(null);
            resolve();
          },
          error: () => { setPendingRpt(null); resolve(); }
        });
      });
    } catch { setPendingRpt(null); }
  };

  const removeReport = (id) => setUploadedReports(prev => { const n={...prev}; delete n[id]; return n; });

  const doAnalyzeReports = async () => {
    const entries = Object.entries(uploadedReports);
    if (!entries.length) return;
    setAnalyzingReports(true);
    const reportsText = entries.map(([,r])=>r.dataText).join('\n\n');
    try {
      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:1000,
          system:`You are a data analyst at BTR Media. Analyze ONLY the provided data. NEVER fabricate numbers. If a metric is not in the data, say "not available in uploaded reports." Return ONLY valid JSON.`,
          messages:[{ role:'user', content:`Analyze this Amazon advertising data for ${client?.name}. Discovery context: ${JSON.stringify(analysis||{})}.

REPORTS:
${reportsText}

Return JSON:
{
  "uploadedReports": ["list of report names analyzed"],
  "keyMetrics": [{ "metric": "name", "value": "exact value from data", "source": "report name", "context": "brief meaning" }],
  "strengthsFound": ["specific strength with exact figure and source report"],
  "gapsFound": ["specific gap with exact figure and source report"],
  "opportunities": ["data-backed opportunity with source"],
  "competitiveInsights": ["any competitive data found, or not available"],
  "recommendedFocus": ["top 3 priorities based purely on data"],
  "dataLimitations": ["what cannot be assessed because of missing reports"]
}` }]
        })
      });
      const d = await res.json();
      const parsed = JSON.parse((d.content?.[0]?.text||'{}').replace(/```json|```/g,'').trim());
      setReportAnalysis(parsed);
      setChatHistory([{ role:'assistant', content:`I have analyzed ${entries.length} report${entries.length>1?'s':''} from ${client?.name}'s account: ${entries.map(([,r])=>r.label).join(', ')}.\n\nI can answer specific questions about the data. What would you like to know?` }]);
    } catch {
      setReportAnalysis({ uploadedReports:entries.map(([,r])=>r.label), keyMetrics:[], strengthsFound:[], gapsFound:[], opportunities:[], competitiveInsights:[], recommendedFocus:[], dataLimitations:['Analysis error — please re-run'] });
    }
    setAnalyzingReports(false);
  };

  const sendChat = async (msgOverride) => {
    const msg = (msgOverride || chatInput).trim();
    if (!msg || chatLoading) return;
    setChatInput('');
    setChatHistory(prev => [...prev, { role:'user', content:msg }]);
    setChatLoading(true);
    const reportsText = Object.values(uploadedReports).map(r=>r.dataText).join('\n\n');
    const allMessages = [...chatHistory, { role:'user', content:msg }];
    try {
      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:1000,
          system:`You are a data analyst at BTR Media analyzing Amazon advertising data for ${client?.name}.

CRITICAL RULES:
1. ONLY cite exact numbers and data that appear in the uploaded reports below
2. When citing a figure, always name the source report
3. If asked about something not in the data, say exactly which report would contain that info and that it has not been uploaded
4. Never estimate, project, or fabricate any metric
5. Be concise and specific

UPLOADED REPORT DATA:
${reportsText}

DISCOVERY ANALYSIS CONTEXT:
${JSON.stringify(analysis||{})}`,
          messages: allMessages.map(m=>({ role:m.role, content:m.content }))
        })
      });
      const d = await res.json();
      setChatHistory(prev => [...prev, { role:'assistant', content:d.content?.[0]?.text||'Unable to process.' }]);
    } catch {
      setChatHistory(prev => [...prev, { role:'assistant', content:'Error connecting. Please try again.' }]);
    }
    setChatLoading(false);
  };

  const doGenerate = async () => {
    setGenerating(true);
    const isProspect = client?.type === 'prospect';
    const hasReportData = Object.keys(uploadedReports).length > 0;
    try {
      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:1000,
          system:`You are a senior Amazon advertising strategist at BTR Media. Generate compelling, specific slide content. Return ONLY a valid JSON array, no markdown.`,
          messages:[{ role:'user', content:`Generate a ${isProspect?'sales pitch':sessionType+' review'} deck for ${client?.name}.
Type: ${client?.type} | Goals: ${client?.goals}
Discovery Analysis: ${JSON.stringify(analysis)}
${hasReportData?`Report Analysis: ${JSON.stringify(reportAnalysis)}\nIMPORTANT: When including specific metrics or numbers in slides, ONLY use figures that appear in the report analysis above. Do not fabricate numbers.`:'No report data uploaded — do not include specific metrics.'}

Return JSON array. Each slide: { id, type, title, headline, bullets (array of strings), notes }.
${isProspect?'Slides: cover, executive-summary, current-state, gap-analysis, solution, dsp-strategy, ctv-strategy, timeline, investment, next-steps':'Slides: cover, month-in-review, kpi-performance, whats-working, needs-attention, strategy, next-steps'}
Make content specific to this client. If no report data, use qualitative language only — no placeholder numbers.` }]
        })
      });
      const d = await res.json();
      setSlides(JSON.parse((d.content?.[0]?.text||'[]').replace(/```json|```/g,'').trim()));
      setStep(3);
    } catch {
      setSlides(isProspect ? buildFallbackSlides(client?.name) : buildFallbackMonthly(client?.name));
      setStep(3);
    }
    setGenerating(false);
  };

  const exportPPTX = async () => {
    setExporting(true);
    try {
      const pptx = new PptxGenJS();
      pptx.layout='LAYOUT_WIDE'; pptx.company='BTR Media';
      const ORG='F97316', DRK='0F172A', WHT='FFFFFF', LGT='E2E8F0', GRY='64748B', MID='1E293B';
      slides.forEach((sl,idx) => {
        const pg = pptx.addSlide();
        pg.addShape(pptx.ShapeType.rect,{x:0,y:0,w:'100%',h:'100%',fill:{color:DRK}});
        pg.addShape(pptx.ShapeType.rect,{x:0,y:0,w:0.12,h:'100%',fill:{color:ORG}});
        if (sl.type==='cover') {
          pg.addText('BTR MEDIA',{x:0.35,y:0.35,w:8,h:0.3,fontSize:10,color:ORG,bold:true,fontFace:'Arial'});
          pg.addText(sl.title||'',{x:0.35,y:1.1,w:8.8,h:1.4,fontSize:40,color:WHT,bold:true,fontFace:'Arial'});
          pg.addText(sl.headline||'',{x:0.35,y:2.7,w:8.5,h:0.6,fontSize:18,color:LGT,fontFace:'Arial'});
          pg.addText(new Date().toLocaleDateString('en-US',{month:'long',year:'numeric'}),{x:0.35,y:3.5,w:4,h:0.35,fontSize:12,color:GRY,fontFace:'Arial'});
        } else {
          pg.addText(`${idx+1}`,{x:9.2,y:0.12,w:0.4,h:0.25,fontSize:9,color:GRY,align:'right',fontFace:'Arial'});
          pg.addShape(pptx.ShapeType.rect,{x:0.35,y:0.5,w:9.1,h:0.035,fill:{color:ORG}});
          pg.addText(sl.title||'',{x:0.35,y:0.6,w:9.1,h:0.6,fontSize:22,color:WHT,bold:true,fontFace:'Arial'});
          if (sl.headline) pg.addText(sl.headline,{x:0.35,y:1.28,w:9.1,h:0.5,fontSize:14,color:ORG,bold:true,italic:true,fontFace:'Arial'});
          const bY = sl.headline?1.92:1.5;
          (sl.bullets||[]).forEach((b,bi) => pg.addText(`  •  ${b}`,{x:0.4,y:bY+(bi*0.52),w:8.9,h:0.48,fontSize:13,color:LGT,fontFace:'Arial',valign:'top'}));
          if (sl.notes) {
            pg.addShape(pptx.ShapeType.rect,{x:0,y:4.82,w:'100%',h:0.65,fill:{color:MID}});
            pg.addText(sl.notes,{x:0.35,y:4.87,w:9.2,h:0.55,fontSize:9,color:GRY,italic:true,fontFace:'Arial'});
          }
          pg.addText('BTR MEDIA — CONFIDENTIAL',{x:0.35,y:5.4,w:5,h:0.2,fontSize:7,color:'334155',fontFace:'Arial'});
        }
      });
      await pptx.writeFile({fileName:`BTR_${(client?.name||'deck').replace(/\s+/g,'_')}_${new Date().toISOString().split('T')[0]}.pptx`});
    } catch(err){ console.error(err); }
    setExporting(false);
  };

  const updateSlide = (id,k,v) => setSlides(p=>p.map(s=>s.id===id?{...s,[k]:v}:s));
  const updateBullet = (sid,bi,v) => setSlides(p=>p.map(s=>{ if(s.id!==sid) return s; const b=[...(s.bullets||[])]; b[bi]=v; return {...s,bullets:b}; }));
  const addBullet = (sid) => setSlides(p=>p.map(s=>s.id===sid?{...s,bullets:[...(s.bullets||[]),'New point']}:s));
  const removeBullet = (sid,bi) => setSlides(p=>p.map(s=>{ if(s.id!==sid) return s; const b=[...(s.bullets||[])]; b.splice(bi,1); return {...s,bullets:b}; }));

  const saveAndAdvance = () => {
    const sess = { id:`s${Date.now()}`, type:sessionType, date:new Date().toISOString(), title:slides[0]?.headline||`${client?.name} Deck`, slideCount:slides.length };
    setClients(p=>p.map(c=>c.id===selId?{...c,sessions:[sess,...c.sessions]}:c));
    setStep(4);
  };

  const startSession = (type) => {
    setSessionType(type); setStep(0); setTranscript(''); setFileName('');
    setAnalysis(null); setUploadedReports({}); setReportAnalysis(null);
    setChatHistory([]); setSlides([]); setEditId(null); setScreen('session');
  };

  const steps = ['Transcript','Discovery','Reports','Edit Deck','Export'];

  const StepBar = () => (
    <div style={{ display:'flex', marginBottom:24 }}>
      {steps.map((s,i)=>(
        <div key={s} style={{ flex:1, padding:'9px 0', textAlign:'center', fontSize:12, fontWeight:600,
          borderBottom:`2px solid ${i<step?C.success:i===step?C.accent:C.border}`,
          color:i<step?C.success:i===step?C.accent:C.faint }}>
          {i<step?'✓ ':''}{s}
        </div>
      ))}
    </div>
  );

  const Clients = () => (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:24 }}>
        <div style={{ fontSize:13, color:C.muted }}>{clients.length} clients · {clients.filter(c=>c.type==='prospect').length} prospects · {clients.filter(c=>c.type==='active').length} active</div>
        <button style={btn('primary')} onClick={()=>setScreen('new-client')}>+ New Client</button>
      </div>
      <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(265px,1fr))', gap:16 }}>
        {clients.map(c=>(
          <div key={c.id} style={{ ...st.card, cursor:'pointer' }} onClick={()=>{ setSelId(c.id); setScreen('profile'); }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:10 }}>
              <div style={{ fontSize:15, fontWeight:700, flex:1, marginRight:8 }}>{c.name}</div>
              <span style={badge(c.type)}>{c.type==='prospect'?'Prospect':'Active'}</span>
            </div>
            <div style={{ fontSize:13, color:C.muted, marginBottom:6 }}>👤 {c.contact}</div>
            <div style={{ fontSize:12, color:C.faint, marginBottom:14, lineHeight:1.5 }}>{c.categories}</div>
            <div style={{ borderTop:`1px solid ${C.border}`, paddingTop:12, display:'flex', justifyContent:'space-between', fontSize:12 }}>
              <span style={{ color:C.faint }}>🗂 {c.sessions.length} deck{c.sessions.length!==1?'s':''}</span>
              <span style={{ color:C.accent, fontWeight:600 }}>Open →</span>
            </div>
          </div>
        ))}
        <div style={{ ...st.card, border:`1px dashed ${C.border}`, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:8, color:C.faint, fontSize:14, minHeight:150 }} onClick={()=>setScreen('new-client')}>
          <span style={{ fontSize:22 }}>＋</span> New Client
        </div>
      </div>
    </div>
  );

  const NewClient = () => {
    const [f,setF] = useState({ name:'', type:'prospect', contact:'', email:'', categories:'', goals:'' });
    const save = () => {
      if (!f.name.trim()) return;
      const nc={...f,id:`c${Date.now()}`,sessions:[],createdAt:new Date().toISOString()};
      setClients(p=>[...p,nc]); setSelId(nc.id); setScreen('profile');
    };
    return (
      <div style={{ maxWidth:580 }}>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>setScreen('clients')}>← Back</button>
        <div style={st.card}>
          <div style={{ fontSize:16, fontWeight:700, marginBottom:20 }}>New Client Profile</div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14 }}>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Company Name *</label><input style={st.input} value={f.name} onChange={e=>setF(p=>({...p,name:e.target.value}))} placeholder="Brand name" /></div>
            <div><label style={st.label}>Client Type</label><select style={st.input} value={f.type} onChange={e=>setF(p=>({...p,type:e.target.value}))}><option value="prospect">Prospect</option><option value="active">Active Client</option></select></div>
            <div><label style={st.label}>Primary Contact</label><input style={st.input} value={f.contact} onChange={e=>setF(p=>({...p,contact:e.target.value}))} placeholder="Name" /></div>
            <div><label style={st.label}>Email</label><input style={st.input} value={f.email} onChange={e=>setF(p=>({...p,email:e.target.value}))} placeholder="email@brand.com" /></div>
            <div><label style={st.label}>Amazon Categories</label><input style={st.input} value={f.categories} onChange={e=>setF(p=>({...p,categories:e.target.value}))} placeholder="Health, Beauty…" /></div>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Goals</label><textarea style={{ ...st.textarea, minHeight:80 }} value={f.goals} onChange={e=>setF(p=>({...p,goals:e.target.value}))} placeholder="What are they trying to achieve on Amazon?" /></div>
          </div>
          <div style={{ display:'flex', gap:10, marginTop:18 }}>
            <button style={{ ...btn('primary'), opacity:f.name?1:0.5 }} onClick={save} disabled={!f.name}>Create Profile</button>
            <button style={btn('ghost')} onClick={()=>setScreen('clients')}>Cancel</button>
          </div>
        </div>
      </div>
    );
  };

  const Profile = () => {
    if (!client) return null;
    return (
      <div>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>setScreen('clients')}>← All Clients</button>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 290px', gap:20 }}>
          <div>
            <div style={st.card}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:14 }}>
                <div>
                  <div style={{ fontSize:22, fontWeight:800 }}>{client.name}</div>
                  <div style={{ marginTop:6, display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
                    <span style={badge(client.type)}>{client.type==='prospect'?'Prospect':'Active Client'}</span>
                    <span style={{ fontSize:12, color:C.faint }}>👤 {client.contact}</span>
                  </div>
                </div>
              </div>
              <div style={{ background:C.surface, borderRadius:8, padding:'10px 14px', marginBottom:12 }}>
                <div style={st.label}>Amazon Categories</div>
                <div style={{ fontSize:13, color:C.text }}>{client.categories||'—'}</div>
              </div>
              <div style={{ background:C.surface, borderRadius:8, padding:'10px 14px' }}>
                <div style={st.label}>Client Goals</div>
                <div style={{ fontSize:13, color:C.text, lineHeight:1.6 }}>{client.goals||'No goals set.'}</div>
              </div>
            </div>
            <div style={{ ...st.card, marginTop:16 }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:14 }}>
                <div style={{ fontSize:15, fontWeight:700 }}>Deck History</div>
                <span style={{ fontSize:12, color:C.faint }}>{client.sessions.length} total</span>
              </div>
              {!client.sessions.length
                ? <div style={{ textAlign:'center', padding:'20px 0', color:C.faint, fontSize:13 }}>No decks yet. Build your first one.</div>
                : client.sessions.map(s=>(
                  <div key={s.id} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'11px 0', borderBottom:`1px solid ${C.border}` }}>
                    <div>
                      <div style={{ fontSize:14, fontWeight:600 }}>{s.title}</div>
                      <div style={{ fontSize:12, color:C.faint, marginTop:2 }}>{new Date(s.date).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})} · {s.slideCount} slides</div>
                    </div>
                    <span style={{ fontSize:11, padding:'3px 9px', borderRadius:20, background:'#374151', color:C.muted }}>{s.type}</span>
                  </div>
                ))
              }
            </div>
          </div>
          <div style={st.card}>
            <div style={{ fontSize:14, fontWeight:700, marginBottom:14 }}>Build a Deck</div>
            {client.type==='prospect'
              ? <button style={{ ...btn('primary'), width:'100%', marginBottom:10 }} onClick={()=>startSession('discovery')}>🎯 New Sales Deck</button>
              : (<>
                  <button style={{ ...btn('primary'), width:'100%', marginBottom:10 }} onClick={()=>startSession('monthly')}>📅 Monthly Review</button>
                  <button style={{ ...btn('ghost'), width:'100%', marginBottom:10 }} onClick={()=>startSession('quarterly')}>📊 Quarterly Review</button>
                  <button style={{ ...btn('ghost'), width:'100%' }} onClick={()=>startSession('annual')}>🏆 Annual Recap</button>
                </>)
            }
            <div style={{ marginTop:14, padding:12, background:C.surface, borderRadius:8, fontSize:12, color:C.faint, lineHeight:1.6 }}>Each session saves to deck history to track progress over time.</div>
          </div>
        </div>
      </div>
    );
  };

  const Session = () => {
    if (!client) return null;
    const uploadedCount = Object.keys(uploadedReports).length;
    const catGroups = REPORT_TYPES.reduce((acc,r)=>{ (acc[r.cat]=acc[r.cat]||[]).push(r); return acc; },{});

    if (step===0) return (
      <div style={{ maxWidth:580 }}>
        <StepBar />
        <div style={st.card}>
          <div style={{ fontSize:16, fontWeight:700, marginBottom:6 }}>Upload Discovery Call Transcript</div>
          <div style={{ fontSize:13, color:C.muted, marginBottom:18 }}>Upload a .txt or .docx transcript. AI will extract pain points, goals, gaps, and key signals.</div>
          <div style={{ border:`2px dashed ${C.border}`, borderRadius:10, padding:28, textAlign:'center', cursor:'pointer', marginBottom:14 }} onClick={()=>transcriptRef.current?.click()}>
            <div style={{ fontSize:36, marginBottom:8 }}>{fileName?'📄':'⬆️'}</div>
            <div style={{ fontSize:14, fontWeight:600, color:fileName?C.success:C.text }}>{fileName||'Click to upload'}</div>
            <div style={{ fontSize:12, color:C.faint, marginTop:4 }}>.txt or .docx supported</div>
            <input ref={transcriptRef} type="file" accept=".txt,.doc,.docx" style={{ display:'none' }} onChange={handleFile} />
          </div>
          <textarea style={{ ...st.textarea, minHeight:140 }} placeholder="Or paste transcript text directly…" value={transcript} onChange={e=>setTranscript(e.target.value)} />
          {transcript && <div style={{ display:'flex', gap:8, padding:'10px 14px', background:'#0F2A1D', border:'1px solid #059669', borderRadius:8, alignItems:'center', margin:'12px 0' }}><span>✅</span><span style={{ fontSize:13, color:C.success }}>{transcript.split(/\s+/).length.toLocaleString()} words ready</span></div>}
          <button style={{ ...btn('primary'), opacity:(!transcript||analyzing)?0.5:1, marginTop:4 }} onClick={doAnalyze} disabled={!transcript||analyzing}>
            {analyzing?'⏳ Analyzing…':'🔍 Analyze with AI'}
          </button>
        </div>
      </div>
    );

    if (step===1 && analysis) return (
      <div>
        <StepBar />
        {analysis.summary && (
          <div style={{ ...st.card, borderLeft:`3px solid ${C.accent}`, marginBottom:16 }}>
            <div style={{ fontSize:11, fontWeight:700, color:C.accent, marginBottom:8, textTransform:'uppercase', letterSpacing:'0.6px' }}>Discovery Brief</div>
            <div style={{ fontSize:14, color:'#E2E8F0', lineHeight:1.7 }}>{analysis.summary}</div>
          </div>
        )}
        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(270px,1fr))', gap:14, marginBottom:20 }}>
          {[['painPoints','Pain Points','🔴','#EF4444'],['goals','Goals','🎯',C.accent],['gaps','Strategy Gaps','⚠️',C.warn],['quickWins','Quick Wins','⚡',C.success],['decisionFactors','Decision Factors','🧠','#A78BFA'],['budgetTimeline','Budget & Timeline','💰',C.blue]].map(([k,l,ic,col])=>(
            <div key={k} style={{ ...st.card, borderTop:`3px solid ${col}` }}>
              <div style={{ fontSize:11, fontWeight:700, color:col, marginBottom:10, textTransform:'uppercase', letterSpacing:'0.5px' }}>{ic} {l}</div>
              {(analysis[k]||[]).map((item,i)=><div key={i} style={{ fontSize:13, color:'#CBD5E1', marginBottom:7, paddingLeft:10, borderLeft:`2px solid ${col}30`, lineHeight:1.5 }}>{item}</div>)}
            </div>
          ))}
        </div>
        <div style={{ display:'flex', gap:10 }}>
          <button style={btn('primary')} onClick={()=>setStep(2)}>Upload Reports →</button>
          <button style={btn('ghost')} onClick={()=>setStep(0)}>← Re-upload</button>
        </div>
      </div>
    );

    if (step===2) return (
      <div>
        <StepBar />
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:20 }}>
          <div>
            <div style={{ fontSize:16, fontWeight:700 }}>Upload Account Reports</div>
            <div style={{ fontSize:13, color:C.muted, marginTop:4 }}>All optional. AI only references data you upload — no numbers are assumed or fabricated.</div>
          </div>
          <div style={{ display:'flex', gap:8, alignItems:'center', flexShrink:0 }}>
            {uploadedCount>0 && <span style={{ fontSize:12, color:C.success, fontWeight:600 }}>✓ {uploadedCount} uploaded</span>}
            <button style={{ ...btn('ghost'), fontSize:12 }} onClick={()=>{ doGenerate(); }}>Skip → Build Deck</button>
          </div>
        </div>
        <input ref={reportFileRef} type="file" accept=".csv,.txt,.tsv" style={{ display:'none' }} onChange={handleReportFile} />

        {Object.entries(catGroups).map(([cat, rpts])=>(
          <div key={cat} style={{ marginBottom:20 }}>
            <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:10 }}>
              <span style={{ width:10, height:10, borderRadius:'50%', background:CAT_COLORS[cat]||C.muted, display:'inline-block' }}></span>
              <span style={{ fontSize:13, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'0.6px' }}>{cat}</span>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(240px,1fr))', gap:10 }}>
              {rpts.map(r=>{
                const up = uploadedReports[r.id];
                return (
                  <div key={r.id} style={{ background:C.card, border:`1px solid ${up?C.success:C.border}`, borderRadius:10, padding:'12px 14px', display:'flex', flexDirection:'column', gap:8 }}>
                    <div style={{ fontSize:13, fontWeight:600, color:up?C.success:C.text }}>{up&&'✓ '}{r.label}</div>
                    {up ? (
                      <div>
                        <div style={{ fontSize:11, color:C.faint }}>{up.fileName} · {up.rowCount.toLocaleString()} rows</div>
                        <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted, marginTop:6, padding:'3px 10px' }} onClick={()=>removeReport(r.id)}>Remove</button>
                      </div>
                    ) : (
                      <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted }} onClick={()=>triggerReportUpload(r.id)}>⬆ Upload CSV</button>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        ))}

        {uploadedCount > 0 && !reportAnalysis && (
          <div style={{ ...st.card, borderTop:`3px solid ${C.accent}`, marginTop:8 }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
              <div>
                <div style={{ fontSize:15, fontWeight:700 }}>Ready to Analyze</div>
                <div style={{ fontSize:13, color:C.muted, marginTop:3 }}>{uploadedCount} report{uploadedCount>1?'s':''} uploaded — AI will extract data-backed insights.</div>
              </div>
              <button style={{ ...btn('primary'), opacity:analyzingReports?0.5:1, flexShrink:0, marginLeft:16 }} onClick={doAnalyzeReports} disabled={analyzingReports}>
                {analyzingReports?'⏳ Analyzing…':'🔍 Analyze Reports'}
              </button>
            </div>
          </div>
        )}

        {reportAnalysis && (
          <div style={{ marginTop:20 }}>
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14, marginBottom:20 }}>
              {[['keyMetrics','Key Metrics','📊',C.blue],['strengthsFound','Strengths Found','✅',C.success],['gapsFound','Gaps Found','⚠️',C.warn],['opportunities','Opportunities','⚡',C.accent],['competitiveInsights','Competitive Insights','🔍','#A78BFA'],['dataLimitations','Data Limitations','🔒',C.faint]].map(([k,l,ic,col])=>(
                reportAnalysis[k]?.length > 0 && (
                  <div key={k} style={{ ...st.card, borderTop:`3px solid ${col}` }}>
                    <div style={{ fontSize:11, fontWeight:700, color:col, marginBottom:10, textTransform:'uppercase', letterSpacing:'0.5px' }}>{ic} {l}</div>
                    {(reportAnalysis[k]||[]).map((item,i)=>{
                      const isMetric = k==='keyMetrics';
                      return isMetric
                        ? <div key={i} style={{ marginBottom:10, padding:'8px 10px', background:C.surface, borderRadius:6 }}>
                            <div style={{ fontSize:13, fontWeight:700, color:C.text }}>{item.metric}: <span style={{ color:col }}>{item.value}</span></div>
                            <div style={{ fontSize:11, color:C.faint, marginTop:2 }}>{item.source} · {item.context}</div>
                          </div>
                        : <div key={i} style={{ fontSize:13, color:'#CBD5E1', marginBottom:7, paddingLeft:10, borderLeft:`2px solid ${col}30`, lineHeight:1.5 }}>{item}</div>;
                    })}
                  </div>
                )
              ))}
            </div>

            <div style={{ ...st.card, borderTop:`3px solid ${C.blue}` }}>
              <div style={{ fontSize:14, fontWeight:700, marginBottom:4 }}>💬 Ask Questions About the Data</div>
              <div style={{ fontSize:12, color:C.muted, marginBottom:14 }}>Answers are strictly based on uploaded report data. No numbers are inferred or fabricated.</div>
              <div style={{ background:C.surface, borderRadius:10, padding:16, minHeight:200, maxHeight:320, overflowY:'auto', marginBottom:12, display:'flex', flexDirection:'column', gap:10 }}>
                {chatHistory.map((m,i)=>(
                  <div key={i} style={{ display:'flex', justifyContent:m.role==='user'?'flex-end':'flex-start' }}>
                    <div style={{ maxWidth:'82%', padding:'10px 14px', borderRadius:10, fontSize:13, lineHeight:1.6, whiteSpace:'pre-wrap',
                      background:m.role==='user'?C.accent:C.card, color:'#fff',
                      borderBottomRightRadius:m.role==='user'?2:10, borderBottomLeftRadius:m.role==='user'?10:2 }}>
                      {m.content}
                    </div>
                  </div>
                ))}
                {chatLoading && <div style={{ display:'flex', justifyContent:'flex-start' }}><div style={{ padding:'10px 14px', background:C.card, borderRadius:10, fontSize:13, color:C.muted }}>Analyzing data…</div></div>}
                <div ref={chatEndRef} />
              </div>
              <div style={{ display:'flex', flexWrap:'wrap', gap:6, marginBottom:10 }}>
                {["What's the top performing search term by ROAS?","Which ASINs are underperforming?","Where are we losing share to competitors?","What's total ad spend vs. revenue?","Which campaigns have the highest TACoS?"].map(q=>(
                  <button key={q} style={{ fontSize:11, padding:'4px 10px', borderRadius:20, border:`1px solid ${C.border}`, background:'transparent', color:C.muted, cursor:'pointer' }} onClick={()=>sendChat(q)}>{q}</button>
                ))}
              </div>
              <div style={{ display:'flex', gap:10 }}>
                <input style={{ ...st.input, flex:1 }} placeholder="Ask anything about the uploaded data…" value={chatInput} onChange={e=>setChatInput(e.target.value)} onKeyDown={e=>e.key==='Enter'&&!e.shiftKey&&sendChat()} />
                <button style={{ ...btn('primary'), padding:'10px 18px', opacity:(!chatInput.trim()||chatLoading)?0.5:1 }} onClick={()=>sendChat()} disabled={!chatInput.trim()||chatLoading}>Send</button>
              </div>
            </div>

            <div style={{ display:'flex', gap:10, marginTop:20 }}>
              <button style={{ ...btn('primary'), opacity:generating?0.5:1 }} onClick={doGenerate} disabled={generating}>
                {generating?'⏳ Generating Deck…':'✨ Generate Deck with Report Data'}
              </button>
            </div>
          </div>
        )}

        {uploadedCount===0 && (
          <div style={{ display:'flex', gap:10, marginTop:8 }}>
            <button style={{ ...btn('primary'), opacity:generating?0.5:1 }} onClick={doGenerate} disabled={generating}>
              {generating?'⏳ Building Deck…':'✨ Build Deck (Discovery Only)'}
            </button>
          </div>
        )}
      </div>
    );

    if (step===3 && slides.length>0) return (
      <div>
        <StepBar />
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:18 }}>
          <div style={{ fontSize:13, color:C.muted }}>{slides.length} slides · click any slide to edit</div>
          <div style={{ display:'flex', gap:10 }}>
            <button style={btn('ghost')} onClick={()=>setStep(2)}>← Reports</button>
            <button style={btn('primary')} onClick={saveAndAdvance}>Export Deck →</button>
          </div>
        </div>
        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(300px,1fr))', gap:14 }}>
          {slides.map((sl,idx)=>(
            <div key={sl.id} style={{ ...st.card, cursor:'pointer', border:`1px solid ${editId===sl.id?C.accent:C.border}` }} onClick={()=>setEditId(editId===sl.id?null:sl.id)}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:10 }}>
                <div style={{ display:'flex', alignItems:'center', gap:8 }}>
                  <span style={{ background:C.accent, color:'#fff', width:22, height:22, borderRadius:'50%', display:'flex', alignItems:'center', justifyContent:'center', fontSize:11, fontWeight:700, flexShrink:0 }}>{idx+1}</span>
                  <span style={{ fontSize:16 }}>{slideIcons[sl.type]||'📄'}</span>
                </div>
                <span style={{ fontSize:10, color:C.faint, textTransform:'uppercase', letterSpacing:'0.4px' }}>{(sl.type||'').replace(/-/g,' ')}</span>
              </div>
              {editId===sl.id ? (
                <div onClick={e=>e.stopPropagation()}>
                  <div style={{ marginBottom:10 }}><label style={st.label}>Title</label><input style={st.input} value={sl.title||''} onChange={e=>updateSlide(sl.id,'title',e.target.value)} /></div>
                  <div style={{ marginBottom:10 }}><label style={st.label}>Headline</label><textarea style={{ ...st.textarea, minHeight:52 }} value={sl.headline||''} onChange={e=>updateSlide(sl.id,'headline',e.target.value)} /></div>
                  <div style={{ marginBottom:8 }}>
                    <label style={st.label}>Bullets</label>
                    {(sl.bullets||[]).map((b,bi)=>(
                      <div key={bi} style={{ display:'flex', gap:6, marginBottom:6 }}>
                        <textarea style={{ ...st.textarea, minHeight:42, fontSize:12, flex:1 }} value={b} onChange={e=>updateBullet(sl.id,bi,e.target.value)} />
                        <button style={{ ...btn('ghost'), padding:'4px 8px', fontSize:12 }} onClick={()=>removeBullet(sl.id,bi)}>✕</button>
                      </div>
                    ))}
                    <button style={{ ...btn('ghost'), fontSize:12, padding:'4px 12px' }} onClick={()=>addBullet(sl.id)}>+ Add bullet</button>
                  </div>
                  <div><label style={st.label}>Speaker Notes</label><textarea style={{ ...st.textarea, minHeight:52, fontSize:12 }} value={sl.notes||''} onChange={e=>updateSlide(sl.id,'notes',e.target.value)} /></div>
                </div>
              ) : (
                <>
                  <div style={{ fontSize:14, fontWeight:700, marginBottom:5 }}>{sl.title}</div>
                  <div style={{ fontSize:12, color:C.accent, fontStyle:'italic', marginBottom:8 }}>{sl.headline}</div>
                  {(sl.bullets||[]).slice(0,3).map((b,bi)=><div key={bi} style={{ fontSize:12, color:C.muted, marginBottom:4, paddingLeft:12, position:'relative' }}><span style={{ position:'absolute', left:0, color:C.accent }}>•</span>{b}</div>)}
                  {(sl.bullets||[]).length>3&&<div style={{ fontSize:11, color:C.faint }}>+{sl.bullets.length-3} more</div>}
                </>
              )}
            </div>
          ))}
        </div>
      </div>
    );

    if (step===4) return (
      <div style={{ maxWidth:520 }}>
        <StepBar />
        <div style={st.card}>
          <div style={{ textAlign:'center', padding:'10px 0 20px' }}>
            <div style={{ fontSize:44, marginBottom:10 }}>🎉</div>
            <div style={{ fontSize:20, fontWeight:800, marginBottom:6 }}>Deck Ready</div>
            <div style={{ fontSize:14, color:C.muted, marginBottom:22 }}>{slides.length} slides · saved to {client?.name}'s history</div>
            <button style={{ ...btn('primary'), width:'100%', padding:'13px 20px', fontSize:15, marginBottom:10 }} onClick={exportPPTX} disabled={exporting}>
              {exporting?'⏳ Exporting…':'⬇️ Download PowerPoint (.pptx)'}
            </button>
          </div>
          <div style={{ borderTop:`1px solid ${C.border}`, paddingTop:16 }}>
            <div style={{ fontSize:13, fontWeight:700, marginBottom:10 }}>Slide List</div>
            {slides.map((sl,i)=>(
              <div key={sl.id} style={{ display:'flex', gap:10, alignItems:'center', padding:'6px 0', borderBottom:'1px solid #1F2937' }}>
                <span style={{ fontSize:11, color:C.faint, width:18 }}>{i+1}</span>
                <span style={{ fontSize:14 }}>{slideIcons[sl.type]||'📄'}</span>
                <span style={{ fontSize:13, color:'#CBD5E1' }}>{sl.title}</span>
              </div>
            ))}
          </div>
          <div style={{ display:'flex', gap:10, marginTop:18 }}>
            <button style={btn('ghost')} onClick={()=>setStep(3)}>← Edit Slides</button>
            <button style={btn('ghost')} onClick={()=>setScreen('profile')}>Back to Profile</button>
          </div>
        </div>
      </div>
    );

    return null;
  };

  const navItems = [
    { id:'clients', label:'All Clients', icon:'👥' },
    ...(selId&&client?[{ id:'profile', label:client.name.split(' ')[0], icon:'📁' }]:[]),
    ...(screen==='new-client'?[{ id:'new-client', label:'New Client', icon:'➕' }]:[]),
    ...(screen==='session'?[{ id:'session', label:'Build Deck', icon:'✨' }]:[]),
  ];
  const meta = {
    clients:{title:'Client Manager',sub:'All prospects and active clients'},
    'new-client':{title:'New Client',sub:'Create a client profile'},
    profile:{title:client?.name||'',sub:`${client?.type==='prospect'?'Prospect':'Active Client'} · ${client?.categories||''}`},
    session:{title:'Deck Builder',sub:`${client?.name||''} · ${sessionType==='discovery'?'Sales Deck':sessionType.charAt(0).toUpperCase()+sessionType.slice(1)+' Review'}`}
  };
  const { title='', sub='' } = meta[screen]||{};

  return (
    <div style={st.app}>
      <div style={st.sidebar}>
        <div style={{ padding:'20px 20px 16px', borderBottom:'1px solid #1F2937' }}>
          <div style={{ fontSize:20, fontWeight:900, color:C.accent, letterSpacing:'-0.5px' }}>BTR</div>
          <div style={{ fontSize:10, color:C.faint, fontWeight:600, letterSpacing:'1.2px', textTransform:'uppercase', marginTop:1 }}>Deck Studio</div>
        </div>
        <div style={{ padding:'12px 10px', flex:1 }}>
          {navItems.map(it=><div key={it.id} style={navItem(screen===it.id)} onClick={()=>setScreen(it.id)}><span>{it.icon}</span><span style={{ overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{it.label}</span></div>)}
        </div>
        <div style={{ padding:'14px 20px', borderTop:'1px solid #1F2937', fontSize:11, color:'#374151' }}>v1.0 · Phase 1</div>
      </div>
      <div style={st.main}>
        <div style={st.header}>
          <div style={{ fontSize:19, fontWeight:700 }}>{title}</div>
          {sub&&<div style={{ fontSize:13, color:C.faint, marginTop:3 }}>{sub}</div>}
        </div>
        <div style={st.content}>
          {screen==='clients'&&<Clients />}
          {screen==='new-client'&&<NewClient />}
          {screen==='profile'&&<Profile />}
          {screen==='session'&&<Session />}
        </div>
      </div>
    </div>
  );
}
