import { useState, useRef, useEffect } from "react";
import * as mammoth from "mammoth";
import Papa from "papaparse";
import PptxGenJS from "pptxgenjs";
import { createClient } from "@supabase/supabase-js";

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

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

const slideIcons = {
  'cover':'🎯','what-we-heard':'👂','account-snapshot':'📸',
  'problem-1':'🔴','problem-2':'🔴','problem-3':'🔴','problem-4':'🔴',
  'cost-of-inaction':'💸','btr-strategy':'🏆','solution-detail-1':'✅',
  'solution-detail-2':'✅','solution-detail-3':'✅','revenue-opportunity':'📈',
  'roadmap':'🗺️','next-steps':'🚀','executive-summary':'📋','current-state':'📊',
  'gap-analysis':'🔍','solution':'💡','dsp-strategy':'📡','ctv-strategy':'📺',
  'timeline':'🗓','investment':'💰','month-in-review':'📅','kpi-performance':'📈',
  'whats-working':'✅','needs-attention':'⚠️','strategy':'🎯',
};

const REPORT_TYPES = [
  { id:'sb_search_term', label:'SB Search Term Report', cat:'Sponsored Brands' },
  { id:'sb_kw_placement', label:'SB Keyword Placement', cat:'Sponsored Brands' },
  { id:'sp_placement', label:'SP Placement Report', cat:'Sponsored Products' },
  { id:'sp_advertised', label:'SP Advertised Product', cat:'Sponsored Products' },
  { id:'sp_search_term', label:'SP Search Term Report', cat:'Sponsored Products' },
  { id:'sqr', label:'SQR (Search Query Report)', cat:'Sponsored Products' },
  { id:'beta_product', label:'Beta Report by Product', cat:'DSP / Beta' },
  { id:'beta_target', label:'Beta Report by Target', cat:'DSP / Beta' },
  { id:'beta_campaign', label:'Beta Report by Campaign', cat:'DSP / Beta' },
  { id:'sc_sales_asin', label:'SC Total Sales by ASIN', cat:'Seller Central' },
  { id:'q1_last_year', label:'Q1 Last Year', cat:'Period Reports' },
  { id:'q1_this_year', label:'Q1 This Year', cat:'Period Reports' },
  { id:'ss_search', label:'SmartScout: Search Terms', cat:'SmartScout' },
  { id:'ss_market', label:'SmartScout: Market Share', cat:'SmartScout' },
  { id:'ss_products', label:'SmartScout: Products', cat:'SmartScout' },
  { id:'ss_brand', label:'SmartScout: Brand Monthly', cat:'SmartScout' },
  { id:'amc_impression_freq', label:'AMC: Impression Frequency', cat:'AMC' },
  { id:'amc_kw_purchase_path', label:'AMC: Sponsored Ads Keyword Purchase Path', cat:'AMC' },
  { id:'amc_ad_placement', label:'AMC: Effectiveness of Ad Placements', cat:'AMC' },
  { id:'amc_time_to_conversion', label:'AMC: Time to Conversion On-Amazon', cat:'AMC' },
];

const CAT_COLORS = { 'Sponsored Brands':'#3B82F6','Sponsored Products':'#8B5CF6','DSP / Beta':'#F97316','Seller Central':'#10B981','Period Reports':'#84CC16','SmartScout':'#F59E0B','AMC':'#EC4899' };

// ── BTR SYSTEM PROMPT (Jess's exact words) ─────────────────────────────────
const BTR_SYSTEM_PROMPT = `You are acting as a senior retail media strategist at BTR Media. I am going to upload performance reports and datasets, discovery call transcripts, and occasionally past marketing plans, audits, or benchmark examples. Your job is to analyze everything thoroughly and help build a full audit and marketing plan for the brand.

When reviewing the data:
- Identify growth opportunities
- Find inefficiencies, wasted spend, structural issues, and strategic gaps
- Look for areas where competitors are outperforming the brand
- Highlight missed opportunities in Amazon Search, DSP, AMC, Walmart, audience strategy, creative strategy, placements, keyword coverage, conquesting, NTB acquisition, and full-funnel execution
- Look for disconnects between the brand's stated goals and what the data actually shows

Use the discovery call transcript to understand: brand priorities, pain points, internal frustrations, revenue goals, team dynamics, risk tolerance, and current strategy limitations.

Do not just summarize data. Build a narrative. The goal is to create an audit and marketing plan that makes the brand clearly understand:
- What they are missing
- Why current strategies are limiting growth
- What opportunities competitors are taking advantage of
- The financial impact of staying where they are
- What becomes possible with BTR Media's approach

The audit should create urgency without sounding overly salesy. It should feel strategic, data-backed, and highly customized to the brand.

Important rules:
- Never invent numbers or performance data
- If information is missing, say so explicitly — do not fabricate
- Challenge assumptions if the data does not support them
- Prioritize insights over fluff
- Think like a consultant, not a generic AI assistant
- Tie findings back to business impact whenever possible
- Focus heavily on storytelling, positioning, and strategic clarity

When building recommendations:
- Prioritize by impact
- Explain the "why" behind each recommendation
- Show how tactics connect together across the funnel
- Make recommendations feel actionable and realistic

Your job is not just to analyze data. Your job is to help build a marketing plan that makes the client feel the cost of inaction and the opportunity in front of them.`;

// ── BTR PLAYBOOK ────────────────────────────────────────────────────────────
const BTR_PLAYBOOK = `
BTR MEDIA STRATEGY PLAYBOOK:

Low Conversion Rate on SP → Isolate highest-converting search terms into single-keyword, single-ASIN ranking campaigns. Prioritize Top of Search placement on proven converting terms.

High TACoS with Flat/Declining Revenue → Rebuild upper-funnel visibility and defend core keyword positions. Declining impression share signals weakening organic rank and over-reliance on bottom-funnel demand capture.

Low New-to-Brand % → Account is over-indexed on conversion tactics. Expand reach through SB, SBV, DSP, and prospecting audiences to introduce the brand earlier in the purchase journey.

Poor Share of Voice vs. Competitors → Shift to exact-match ranking campaigns on highest-value search terms. Own premium placements that drive both conversions and organic rank growth.

No Mid-Funnel Coverage → Increase SB, SBV, and audience-based campaigns to keep the brand visible while shoppers evaluate options.

High Impression Frequency with Low Purchase Rate (AMC) → Build sequential remarketing through DSP and Sponsored Ads. Conversion rates improve significantly as shoppers accumulate multiple meaningful brand interactions.

Long Time-to-Conversion (AMC) → Align attribution windows, remarketing durations, and messaging with the brand's actual purchase cycle.

Competitors Conquesting Your ASINs → Protect core branded searches and high-value ASINs while actively conquesting competitor traffic for NTB acquisition.`;

// ── REPORT DIAGNOSTICS ─────────────────────────────────────────────────────
const REPORT_DIAGNOSTICS = {
  sp_search_term:`SP SEARCH TERM — look for: high spend/low CVR (<5%) terms, strong ROAS terms with low impression share, branded terms in broad/auto, single term driving >30% of spend, zero-order terms for negation, impression share gaps vs. available volume.`,
  sb_kw_placement:`SB KEYWORD PLACEMENT — look for: TOS % below 40% on branded terms, high spend on rest-of-search with low CTR, keywords winning placements but not converting.`,
  sp_placement:`SP PLACEMENT — look for: TOS ROAS vs. Rest of Search ROAS gap, product page placements with low ROAS, bid modifiers misaligned with performance.`,
  sp_advertised:`SP ADVERTISED PRODUCT — look for: ASINs with high spend below 2x ROAS, strong ROAS ASINs with low impression share, revenue concentration risk, zero-order ASINs after significant spend.`,
  sqr:`SQR — look for: share of voice on branded vs. non-branded, competitor branded terms in query mix, high-volume queries with low brand impression share.`,
  beta_product:`DSP BETA BY PRODUCT — look for: ROAS below 1.5x, high frequency with low DPVR, strong view rate but low purchase rate.`,
  beta_target:`DSP BETA BY TARGET — look for: high frequency audiences with low conversion, retargeting vs. prospecting spend imbalance.`,
  beta_campaign:`DSP BETA BY CAMPAIGN — look for: ROAS trends, underspending campaigns, awareness vs. retargeting budget balance.`,
  sc_sales_asin:`SC TOTAL SALES BY ASIN — look for: revenue concentration, declining ASIN trends, organic vs. paid revenue ratio.`,
  q1_last_year:`Q1 LAST YEAR — use as YoY baseline for revenue, ROAS, NTB, category rank.`,
  q1_this_year:`Q1 THIS YEAR — compare to Q1 last year: revenue growth rate vs. spend growth, ROAS direction, market share movement.`,
  ss_search:`SMARTSCOUT SEARCH TERMS — look for: high-volume terms where brand has <10% impression share, terms competitors dominate, brand terms being captured by competitors, emerging terms.`,
  ss_market:`SMARTSCOUT MARKET SHARE — look for: brands gaining vs. losing share, gap between this brand and top 3 competitors, share trend acceleration.`,
  ss_products:`SMARTSCOUT PRODUCTS — look for: competitor ASINs outranking brand, review/rating gaps, price positioning, BSR trends.`,
  ss_brand:`SMARTSCOUT BRAND MONTHLY — look for: revenue trend vs. category trend, estimated ad spend vs. organic revenue, growth rate vs. top 3 competitors.`,
  amc_impression_freq:`AMC IMPRESSION FREQUENCY — look for: average frequency above 8 with low conversion, frequency distribution, large audience at 1-2 exposures not converting.`,
  amc_kw_purchase_path:`AMC KEYWORD PURCHASE PATH — look for: touchpoints before conversion, drop-off points, assisted vs. last-touch keyword performance.`,
  amc_ad_placement:`AMC AD PLACEMENTS — look for: which ad types drive most assisted conversions, placement efficiency gaps, cross-channel sequence gaps.`,
  amc_time_to_conversion:`AMC TIME TO CONVERSION — look for: average days to conversion, % converting after 7 days, attribution window misalignment with purchase cycle.`,
};

const FALLBACK_ANALYSIS = {
  painPoints:['Low visibility vs. category competitors','Poor ROAS on Sponsored Products','No DSP or programmatic strategy'],
  goals:['Increase market share','Improve ROAS','Launch DSP retargeting'],
  currentStrategy:['Basic Sponsored Products with manual bidding','No external traffic or DSP'],
  gaps:['No DSP, programmatic, or CTV/STV strategy','Missing mid-funnel coverage'],
  keyMetrics:['Analysis could not connect — using template data'],
  budgetTimeline:['To be confirmed'],
  decisionFactors:['Proven results','Transparent reporting'],
  competitiveContext:['Competitors investing in DSP and programmatic'],
  quickWins:['Launch DSP retargeting','Restructure SP campaigns'],
  summary:'Discovery analysis could not be processed automatically. Please review and edit the sections below, or re-run the analysis.',
};

const buildFallbackSlides = (n) => ([
  { id:'sl-1', type:'cover', title:n, headline:'Retail Media Audit & Growth Proposal', bullets:[], notes:'Open by establishing this is a partnership conversation.' },
  { id:'sl-2', type:'what-we-heard', title:'What We Heard', headline:'We listened before we analyzed.', bullets:['[Primary goal from discovery call]','[Key frustration or challenge stated]','[What they want from an agency partner]','BTR Media was built to solve exactly this.'], notes:'Mirror back their exact words. This builds trust before anything else.' },
  { id:'sl-3', type:'account-snapshot', title:'Account Snapshot', headline:'Here is where the account stands today.', bullets:['[Total ad spend — source: account data]','[Blended ROAS vs. peer benchmark]','[SP Conversion Rate]','[Key performance gap vs. goal]'], notes:'Replace brackets with actual data before presenting.' },
  { id:'sl-4', type:'problem-1', title:'Problem #1: [Name the Problem]', headline:'[Data-backed headline that names the issue directly]', bullets:['[Exact data point proving the problem exists]','[What this is costing the brand in revenue or efficiency]','[Why this is happening — the structural cause]'], notes:'Be direct. The evidence speaks for itself.' },
  { id:'sl-5', type:'problem-2', title:'Problem #2: [Name the Problem]', headline:'[Data-backed headline]', bullets:['[Evidence from data]','[Business impact]','[Root cause]'], notes:'' },
  { id:'sl-6', type:'problem-3', title:'Problem #3: [Name the Problem]', headline:'[Data-backed headline]', bullets:['[Evidence from data]','[Business impact]','[Root cause]'], notes:'' },
  { id:'sl-7', type:'cost-of-inaction', title:'The Cost of Staying Here', headline:'Every month without a fix is revenue left on the table.', bullets:['[Quantify what the current gaps are costing]','[Competitive threat — what competitors are gaining while this brand stands still]','[The compounding effect of structural problems over time]'], notes:'This slide creates urgency through math, not pressure.' },
  { id:'sl-8', type:'btr-strategy', title:'The BTR Media Approach', headline:'Full-funnel strategy built around your specific gaps.', bullets:['Campaign restructure and keyword ownership — solving the foundation','DSP activation — capturing demand being left behind','AMC intelligence — making every dollar work harder','Dedicated AM team with the expertise to execute'], notes:'Connect each element to a problem identified earlier.' },
  { id:'sl-9', type:'solution-detail-1', title:'Priority Fix #1: [Solution Name]', headline:'[Specific BTR recommendation]', bullets:['What we do — specific tactic or structural change','Why this solves the identified problem','What result this drives — tied to their stated goal'], notes:'' },
  { id:'sl-10', type:'solution-detail-2', title:'Priority Fix #2: [Solution Name]', headline:'[Specific BTR recommendation]', bullets:['What we do','Why this solves the problem','What result this drives'], notes:'' },
  { id:'sl-11', type:'revenue-opportunity', title:'What Becomes Possible', headline:'This is what growth looks like when the gaps are closed.', bullets:['[Efficiency improvement — ROAS or TACoS target]','[NTB growth potential from demand generation activation]','[Share of voice recovery on key category terms]','[Revenue range that becomes attainable with full-funnel execution]'], notes:'Only use projections that can be supported by the data. Qualify everything.' },
  { id:'sl-12', type:'roadmap', title:'90-Day Roadmap', headline:'From audit to full-funnel activation.', bullets:['Weeks 1-2: Account audit, campaign restructure, strategy alignment','Weeks 3-4: New campaign builds, DSP setup, tracking dashboards live','Weeks 5-8: Optimize, AMC audience activation, creative testing','Weeks 9-12: Full-funnel scaling, cross-channel expansion'], notes:'' },
  { id:'sl-13', type:'next-steps', title:'Next Steps', headline:'Ready to build this together?', bullets:['Share Amazon Seller/Vendor Central access for full account audit','Send over SmartScout report for competitive analysis','Schedule kickoff call within 5 business days','BTR team ready to activate within 1 week of agreement'], notes:'End with clear, easy actions.' },
]);

const buildFallbackMonthly = (n) => ([
  { id:'sl-1', type:'cover', title:`${n} — Monthly Review`, headline:new Date().toLocaleDateString('en-US',{month:'long',year:'numeric'}), bullets:[], notes:'' },
  { id:'sl-2', type:'what-we-heard', title:'Your Goals This Month', headline:'What we set out to accomplish.', bullets:['[Primary goal for the month]','[KPI targets agreed upon]','[Strategic priority from last session]'], notes:'' },
  { id:'sl-3', type:'account-snapshot', title:'Month in Review', headline:'Key wins, optimizations, and what we moved on.', bullets:['[Top campaign win — replace with actuals]','[Budget efficiency improvement]','[New tests or activations launched]'], notes:'' },
  { id:'sl-4', type:'kpi-performance', title:'KPI Performance vs. Goals', headline:'Tracking toward target.', bullets:['Revenue: $X vs. $Y goal','ROAS: X.Xx vs. Y.Yx target','New to Brand: X%','Share of Voice: X%'], notes:'Pull actuals before presenting.' },
  { id:'sl-5', type:'whats-working', title:"What's Working", headline:'Double down on these.', bullets:['[Top campaign / tactic]','[Best converting audience]','[Top ASIN or keyword]'], notes:'' },
  { id:'sl-6', type:'needs-attention', title:'Needs Attention', headline:'Active optimizations in progress.', bullets:['[Underperforming area — action taken]','[Gap vs. target — close plan]'], notes:'' },
  { id:'sl-7', type:'strategy', title:'Next Month Priorities', headline:'Three things we are focused on.', bullets:['Priority 1','Priority 2','Priority 3'], notes:'' },
  { id:'sl-8', type:'next-steps', title:'Action Items', headline:'What happens next.', bullets:['BTR: Optimizations live by [date]','Client: Assets by [date]','Both: Check-in on [date]'], notes:'' },
]);

export default function BTRDeckStudio() {
  const [screen, setScreen] = useState('clients');
  const [clients, setClients] = useState([]);
  const [loadingClients, setLoadingClients] = useState(true);
  const [selId, setSelId] = useState(null);
  const [sessionType, setSessionType] = useState('discovery');
  const [step, setStep] = useState(0);
  const [transcript, setTranscript] = useState('');
  const [fileName, setFileName] = useState('');
  const [refDeck, setRefDeck] = useState(null);
  const [refDeckName, setRefDeckName] = useState('');
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
  const [savingSession, setSavingSession] = useState(false);
  const transcriptRef = useRef(null);
  const refDeckRef = useRef(null);
  const reportFileRef = useRef(null);
  const chatEndRef = useRef(null);

  useEffect(() => { loadClients(); }, []);
  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior:'smooth' }); }, [chatHistory]);

  const loadClients = async () => {
    setLoadingClients(true);
    const { data } = await supabase.from('clients').select('*, sessions(*)').order('created_at', { ascending:false });
    if (data) setClients(data.map(c => ({ ...c, sessions:(c.sessions||[]).sort((a,b)=>new Date(b.created_at)-new Date(a.created_at)) })));
    setLoadingClients(false);
  };

  const client = clients.find(c => c.id === selId);

  const handleFile = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    setFileName(f.name);
    if (f.name.match(/\.docx?$/i)) { const buf=await f.arrayBuffer(); const r=await mammoth.extractRawText({arrayBuffer:buf}); setTranscript(r.value); }
    else { setTranscript(await f.text()); }
  };

  const handleRefDeck = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    setRefDeckName(f.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      const base64 = ev.target.result.split(',')[1];
      setRefDeck(base64);
    };
    reader.readAsDataURL(f);
  };

  const doAnalyze = async () => {
    setAnalyzing(true);
    try {
      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:1000,
          system:`${BTR_SYSTEM_PROMPT}\n\nAnalyze discovery call transcripts for sales intelligence. Return ONLY valid JSON, no markdown.`,
          messages:[{ role:'user', content:`Analyze this discovery call transcript for ${client?.name}. Return JSON with: painPoints, goals, currentStrategy, gaps, keyMetrics, budgetTimeline, decisionFactors, competitiveContext, quickWins (each an array of strings), and summary (2-3 sentence string).\n\nTRANSCRIPT:\n${transcript}` }]
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
            const rows = allRows.slice(0,75);
            const dataText = [`=== ${rtype.label} (${rtype.cat}) ===`,`Total rows: ${allRows.length} | Showing first ${rows.length}`,`Columns: ${headers.join(', ')}`,'---',rows.map(row=>headers.map(h=>`${h}: ${row[h]??''}`).join(' | ')).join('\n')].join('\n');
            setUploadedReports(prev => ({ ...prev, [pendingRpt]:{ label:rtype.label, cat:rtype.cat, fileName:f.name, rowCount:allRows.length, headers, dataText } }));
            setPendingRpt(null); resolve();
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
    const diagnosticContext = entries.map(([id,r]) => {
      const diag = REPORT_DIAGNOSTICS[id];
      return diag ? `\n${diag}\n\nDATA:\n${r.dataText}` : `\n=== ${r.label} ===\n${r.dataText}`;
    }).join('\n\n---\n\n');
    try {
      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:4000,
          system:`${BTR_SYSTEM_PROMPT}\n\n${BTR_PLAYBOOK}\n\nAnalyze ONLY the provided data using the diagnostic criteria for each report type. NEVER fabricate numbers. Return ONLY valid JSON.`,
          messages:[{ role:'user', content:`Analyze this Amazon advertising data for ${client?.name}.\nDiscovery context: ${JSON.stringify(analysis||{})}\n\n${diagnosticContext}\n\nReturn JSON:\n{\n  "uploadedReports": ["list of reports analyzed"],\n  "keyMetrics": [{ "metric": "name", "value": "exact value", "source": "report name", "context": "brief meaning" }],\n  "strengthsFound": ["specific strength with exact figure and source"],\n  "gapsFound": ["specific gap with exact figure and source"],\n  "opportunities": ["data-backed opportunity with source"],\n  "competitiveInsights": ["competitive data or state not available"],\n  "strategicRecommendations": [{ "problem": "specific problem", "evidence": "exact data points proving it", "btrStrategy": "specific BTR approach", "priority": "high or medium or low", "deckSlide": "slide type this should inform" }],\n  "dataLimitations": ["what cannot be assessed due to missing reports"]\n}` }]
        })
      });
      const d = await res.json();
      const parsed = JSON.parse((d.content?.[0]?.text||'{}').replace(/```json|```/g,'').trim());
      setReportAnalysis(parsed);
      setChatHistory([{ role:'assistant', content:`I have analyzed ${entries.length} report${entries.length>1?'s':''} from ${client?.name}'s account using BTR's diagnostic framework.\n\nI found ${parsed.strategicRecommendations?.length||0} strategic recommendations. What would you like to dig into?` }]);
    } catch(err) {
      console.error(err);
      setReportAnalysis({ uploadedReports:entries.map(([,r])=>r.label), keyMetrics:[], strengthsFound:[], gapsFound:[], opportunities:[], competitiveInsights:[], strategicRecommendations:[], dataLimitations:['Analysis error — please re-run'] });
    }
    setAnalyzingReports(false);
  };

  const sendChat = async (msgOverride) => {
    const msg = (msgOverride||chatInput).trim();
    if (!msg||chatLoading) return;
    setChatInput('');
    setChatHistory(prev=>[...prev,{ role:'user', content:msg }]);
    setChatLoading(true);
    const reportsText = Object.values(uploadedReports).map(r=>r.dataText.split('\n').slice(0,20).join('\n')).join('\n\n');
    try {
      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:2000,
          system:`${BTR_SYSTEM_PROMPT}\n\n${BTR_PLAYBOOK}\n\nYou are analyzing Amazon advertising data for ${client?.name}. ONLY cite exact numbers from the uploaded reports. Name the source report when citing a figure. If asked about something not in the data, name which report would contain it.\n\nREPORT ANALYSIS SUMMARY:\n${JSON.stringify(reportAnalysis||{})}\n\nRAW DATA SAMPLE:\n${reportsText}\n\nDISCOVERY CONTEXT:\n${JSON.stringify(analysis||{})}`,
          messages:[...chatHistory,{ role:'user', content:msg }].map(m=>({ role:m.role, content:m.content }))
        })
      });
      const d = await res.json();
      setChatHistory(prev=>[...prev,{ role:'assistant', content:d.content?.[0]?.text||'Unable to process.' }]);
    } catch {
      setChatHistory(prev=>[...prev,{ role:'assistant', content:'Error connecting. Please try again.' }]);
    }
    setChatLoading(false);
  };

  // ── DECK GENERATION — NARRATIVE AUDIT ENGINE ───────────────────────────────
  const doGenerate = async () => {
    setGenerating(true);
    const isProspect = client?.type === 'prospect';
    const hasReportData = reportAnalysis && (reportAnalysis.strategicRecommendations?.length > 0 || reportAnalysis.keyMetrics?.length > 0);

    try {
      // Build message content array — include reference deck PDF if uploaded
      const msgContent = [];

      if (refDeck) {
        msgContent.push({
          type:'document',
          source:{ type:'base64', media_type:'application/pdf', data:refDeck }
        });
      }

      msgContent.push({
        type:'text',
        text:`Build a complete ${isProspect ? 'retail media audit and sales pitch deck' : sessionType + ' performance review deck'} for ${client?.name}.

DISCOVERY CALL ANALYSIS:
${JSON.stringify(analysis)}

${hasReportData ? `ACCOUNT DATA AND STRATEGIC FINDINGS (use ONLY these numbers — never fabricate):
${JSON.stringify(reportAnalysis)}` : 'No account data uploaded — use qualitative language only. No placeholder numbers.'}

CLIENT GOALS: ${client?.goals || 'See discovery analysis above'}

${refDeck ? `A reference BTR Media deck has been uploaded. Study its narrative structure, depth of insight, data presentation style, and storytelling approach. Build this deck to that same standard of quality and specificity.` : ''}

BUILD THIS AS A NARRATIVE AUDIT — NOT A GENERIC SLIDE DECK.

Structure the deck to tell this story:
1. WHAT WE HEARD — Open by reflecting back exactly what the client told us. Their words. Their goals. Their frustrations. Show them we were listening. Then show them what BTR delivers against each point.
2. ACCOUNT SNAPSHOT — Show the current state of the account with specific data. Key metrics. Where they stand vs. benchmarks. No spin — honest picture.
3. THE PROBLEMS (2-4 problem slides) — For each problem found: name it directly, show the exact data that proves it exists, explain what it is costing them in revenue or efficiency, explain why it is happening structurally.
4. THE COST OF STAYING HERE — One slide that quantifies the financial impact of inaction. What competitors are gaining. What this brand is leaving on the table every month.
5. THE BTR STRATEGY — Full-funnel approach, connected to each identified problem. Not a capabilities slide — a solutions slide.
6. PRIORITY FIXES (2-3 slides) — Specific BTR recommendations, one per slide. What we do, why it solves the problem, what result it drives.
7. WHAT BECOMES POSSIBLE — Revenue and growth potential when the gaps are closed. Only use numbers supportable by the data.
8. 90-DAY ROADMAP — Phase 1 through 4, specific and actionable.
9. NEXT STEPS — Clear, easy actions to move forward.

RULES:
- Every problem slide must cite exact data — no vague claims
- Every solution must map directly to an identified problem
- Create urgency through insight and data, not sales pressure
- Write like a consultant who spent a week in this account, not a template generator
- Be specific to ${client?.name} — zero generic language

Return ONLY a JSON array. Each slide: { id, type, title, headline, bullets (array of specific 1-2 sentence strings), notes (presenter talking points) }

Slide types: cover, what-we-heard, account-snapshot, problem-1, problem-2, problem-3, cost-of-inaction, btr-strategy, solution-detail-1, solution-detail-2, revenue-opportunity, roadmap, next-steps`
      });

      const res = await fetch('/api/ai', {
        method:'POST', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({ model:'claude-sonnet-4-20250514', max_tokens:4000, system:BTR_SYSTEM_PROMPT, messages:[{ role:'user', content:msgContent }] })
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
          const bY=sl.headline?1.92:1.5;
          (sl.bullets||[]).forEach((b,bi)=>pg.addText(`  •  ${b}`,{x:0.4,y:bY+(bi*0.52),w:8.9,h:0.48,fontSize:13,color:LGT,fontFace:'Arial',valign:'top'}));
          if (sl.notes) { pg.addShape(pptx.ShapeType.rect,{x:0,y:4.82,w:'100%',h:0.65,fill:{color:MID}}); pg.addText(sl.notes,{x:0.35,y:4.87,w:9.2,h:0.55,fontSize:9,color:GRY,italic:true,fontFace:'Arial'}); }
          pg.addText('BTR MEDIA — CONFIDENTIAL',{x:0.35,y:5.4,w:5,h:0.2,fontSize:7,color:'334155',fontFace:'Arial'});
        }
      });
      await pptx.writeFile({fileName:`BTR_${(client?.name||'deck').replace(/\s+/g,'_')}_${new Date().toISOString().split('T')[0]}.pptx`});
    } catch(err){ console.error(err); }
    setExporting(false);
  };

  const updateSlide=(id,k,v)=>setSlides(p=>p.map(s=>s.id===id?{...s,[k]:v}:s));
  const updateBullet=(sid,bi,v)=>setSlides(p=>p.map(s=>{ if(s.id!==sid) return s; const b=[...(s.bullets||[])]; b[bi]=v; return {...s,bullets:b}; }));
  const addBullet=(sid)=>setSlides(p=>p.map(s=>s.id===sid?{...s,bullets:[...(s.bullets||[]),'New point']}:s));
  const removeBullet=(sid,bi)=>setSlides(p=>p.map(s=>{ if(s.id!==sid) return s; const b=[...(s.bullets||[])]; b.splice(bi,1); return {...s,bullets:b}; }));

  const saveAndAdvance = async () => {
    setSavingSession(true);
    const sessionData = { client_id:selId, type:sessionType, title:slides[0]?.headline||`${client?.name} Deck`, slide_count:slides.length, slides_json:slides, analysis_json:analysis, report_analysis_json:reportAnalysis };
    const { data } = await supabase.from('sessions').insert([sessionData]).select().single();
    if (data) setClients(prev=>prev.map(c=>c.id===selId?{ ...c, sessions:[data,...(c.sessions||[])] }:c));
    setSavingSession(false);
    setStep(4);
  };

  const startSession = (type) => {
    setSessionType(type); setStep(0); setTranscript(''); setFileName('');
    setRefDeck(null); setRefDeckName('');
    setAnalysis(null); setUploadedReports({}); setReportAnalysis(null);
    setChatHistory([]); setSlides([]); setEditId(null); setScreen('session');
  };

  const steps = ['Transcript','Discovery','Reports','Edit Deck','Export'];
  const StepBar = () => (
    <div style={{ display:'flex', marginBottom:24 }}>
      {steps.map((s,i)=>(
        <div key={s} style={{ flex:1, padding:'9px 0', textAlign:'center', fontSize:12, fontWeight:600, borderBottom:`2px solid ${i<step?C.success:i===step?C.accent:C.border}`, color:i<step?C.success:i===step?C.accent:C.faint }}>
          {i<step?'✓ ':''}{s}
        </div>
      ))}
    </div>
  );

  const Clients = () => (
    <div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:24 }}>
        <div style={{ fontSize:13, color:C.muted }}>{loadingClients?'Loading...':`${clients.length} clients · ${clients.filter(c=>c.type==='prospect').length} prospects · ${clients.filter(c=>c.type==='active').length} active`}</div>
        <button style={btn('primary')} onClick={()=>setScreen('new-client')}>+ New Client</button>
      </div>
      {loadingClients ? <div style={{ textAlign:'center', padding:'60px 0', color:C.faint }}>Loading clients…</div> : (
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
                <span style={{ color:C.faint }}>🗂 {(c.sessions||[]).length} deck{(c.sessions||[]).length!==1?'s':''}</span>
                <span style={{ color:C.accent, fontWeight:600 }}>Open →</span>
              </div>
            </div>
          ))}
          <div style={{ ...st.card, border:`1px dashed ${C.border}`, cursor:'pointer', display:'flex', alignItems:'center', justifyContent:'center', gap:8, color:C.faint, fontSize:14, minHeight:150 }} onClick={()=>setScreen('new-client')}>
            <span style={{ fontSize:22 }}>＋</span> New Client
          </div>
        </div>
      )}
    </div>
  );

  const NewClient = () => {
    const [f,setF]=useState({ name:'', type:'prospect', contact:'', email:'', categories:'', goals:'' });
    const [saving,setSaving]=useState(false);
    const save=async()=>{ if(!f.name.trim()) return; setSaving(true); const{data}=await supabase.from('clients').insert([f]).select().single(); if(data){setClients(p=>[{...data,sessions:[]},...p]);setSelId(data.id);setScreen('profile');} setSaving(false); };
    return (
      <div style={{ maxWidth:580 }}>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>setScreen('clients')}>← Back</button>
        <div style={st.card}>
          <div style={{ fontSize:16, fontWeight:700, marginBottom:20 }}>New Client Profile</div>
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14 }}>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Company Name *</label><input style={st.input} value={f.name} onChange={e=>setF(p=>({...p,name:e.target.value}))} placeholder="Brand name" /></div>
            <div><label style={st.label}>Client Type</label><select style={st.input} value={f.type} onChange={e=>setF(p=>({...p,type:e.target.value}))}><option value="prospect">Prospect</option><option value="active">Active Client</option></select></div>
            <div><label style={st.label}>Contact</label><input style={st.input} value={f.contact} onChange={e=>setF(p=>({...p,contact:e.target.value}))} placeholder="Name" /></div>
            <div><label style={st.label}>Email</label><input style={st.input} value={f.email} onChange={e=>setF(p=>({...p,email:e.target.value}))} placeholder="email@brand.com" /></div>
            <div><label style={st.label}>Amazon Categories</label><input style={st.input} value={f.categories} onChange={e=>setF(p=>({...p,categories:e.target.value}))} placeholder="Health, Beauty…" /></div>
            <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Goals</label><textarea style={{ ...st.textarea, minHeight:80 }} value={f.goals} onChange={e=>setF(p=>({...p,goals:e.target.value}))} placeholder="What are they trying to achieve on Amazon?" /></div>
          </div>
          <div style={{ display:'flex', gap:10, marginTop:18 }}>
            <button style={{ ...btn('primary'), opacity:(!f.name||saving)?0.5:1 }} onClick={save} disabled={!f.name||saving}>{saving?'Saving…':'Create Profile'}</button>
            <button style={btn('ghost')} onClick={()=>setScreen('clients')}>Cancel</button>
          </div>
        </div>
      </div>
    );
  };

  const Profile = () => {
    if (!client) return null;
    const [editing,setEditing]=useState(false);
    const [ef,setEf]=useState({ name:client.name, type:client.type, contact:client.contact||'', email:client.email||'', categories:client.categories||'', goals:client.goals||'' });
    const [saving,setSaving]=useState(false);
    const saveEdits=async()=>{ setSaving(true); const{data}=await supabase.from('clients').update(ef).eq('id',selId).select().single(); if(data) setClients(prev=>prev.map(c=>c.id===selId?{...data,sessions:c.sessions}:c)); setSaving(false); setEditing(false); };
    return (
      <div>
        <button style={{ ...btn('outline'), marginBottom:18 }} onClick={()=>setScreen('clients')}>← All Clients</button>
        <div style={{ display:'grid', gridTemplateColumns:'1fr 290px', gap:20 }}>
          <div>
            <div style={st.card}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:14 }}>
                <div style={{ flex:1, marginRight:12 }}>
                  {editing ? <input style={{ ...st.input, fontSize:20, fontWeight:800, marginBottom:8 }} value={ef.name} onChange={e=>setEf(p=>({...p,name:e.target.value}))} /> : <div style={{ fontSize:22, fontWeight:800 }}>{client.name}</div>}
                  <div style={{ marginTop:6, display:'flex', gap:8, alignItems:'center', flexWrap:'wrap' }}>
                    {editing ? <select style={{ ...st.input, width:'auto', fontSize:12 }} value={ef.type} onChange={e=>setEf(p=>({...p,type:e.target.value}))}><option value="prospect">Prospect</option><option value="active">Active Client</option></select> : <span style={badge(client.type)}>{client.type==='prospect'?'Prospect':'Active Client'}</span>}
                    {!editing && <span style={{ fontSize:12, color:C.faint }}>👤 {client.contact}</span>}
                  </div>
                </div>
                <div style={{ display:'flex', gap:8, flexShrink:0 }}>
                  {editing ? (<><button style={{ ...btn('primary'), opacity:saving?0.5:1 }} onClick={saveEdits} disabled={saving}>{saving?'Saving…':'Save'}</button><button style={btn('ghost')} onClick={()=>{ setEditing(false); setEf({ name:client.name, type:client.type, contact:client.contact||'', email:client.email||'', categories:client.categories||'', goals:client.goals||'' }); }}>Cancel</button></>) : (<button style={btn('ghost')} onClick={()=>setEditing(true)}>✏️ Edit</button>)}
                </div>
              </div>
              {editing ? (
                <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:12 }}>
                  <div><label style={st.label}>Contact</label><input style={st.input} value={ef.contact} onChange={e=>setEf(p=>({...p,contact:e.target.value}))} /></div>
                  <div><label style={st.label}>Email</label><input style={st.input} value={ef.email} onChange={e=>setEf(p=>({...p,email:e.target.value}))} /></div>
                  <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Amazon Categories</label><input style={st.input} value={ef.categories} onChange={e=>setEf(p=>({...p,categories:e.target.value}))} /></div>
                  <div style={{ gridColumn:'1/-1' }}><label style={st.label}>Goals</label><textarea style={{ ...st.textarea, minHeight:100 }} value={ef.goals} onChange={e=>setEf(p=>({...p,goals:e.target.value}))} /></div>
                </div>
              ) : (
                <>
                  <div style={{ background:C.surface, borderRadius:8, padding:'10px 14px', marginBottom:12 }}><div style={st.label}>Contact & Email</div><div style={{ fontSize:13, color:C.text }}>👤 {client.contact||'—'}{client.email?` · ${client.email}`:''}</div></div>
                  <div style={{ background:C.surface, borderRadius:8, padding:'10px 14px', marginBottom:12 }}><div style={st.label}>Amazon Categories</div><div style={{ fontSize:13, color:C.text }}>{client.categories||'—'}</div></div>
                  <div style={{ background:C.surface, borderRadius:8, padding:'10px 14px' }}><div style={st.label}>Goals</div><div style={{ fontSize:13, color:C.text, lineHeight:1.6 }}>{client.goals||'No goals set.'}</div></div>
                </>
              )}
            </div>
            <div style={{ ...st.card, marginTop:16 }}>
              <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:14 }}>
                <div style={{ fontSize:15, fontWeight:700 }}>Deck History</div>
                <span style={{ fontSize:12, color:C.faint }}>{(client.sessions||[]).length} total</span>
              </div>
              {!(client.sessions||[]).length ? <div style={{ textAlign:'center', padding:'20px 0', color:C.faint, fontSize:13 }}>No decks yet.</div> : (client.sessions||[]).map(s=>(
                <div key={s.id} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'11px 0', borderBottom:`1px solid ${C.border}` }}>
                  <div><div style={{ fontSize:14, fontWeight:600 }}>{s.title}</div><div style={{ fontSize:12, color:C.faint, marginTop:2 }}>{new Date(s.created_at).toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'})} · {s.slide_count} slides</div></div>
                  <span style={{ fontSize:11, padding:'3px 9px', borderRadius:20, background:'#374151', color:C.muted }}>{s.type}</span>
                </div>
              ))}
            </div>
          </div>
          <div style={st.card}>
            <div style={{ fontSize:14, fontWeight:700, marginBottom:14 }}>Build a Deck</div>
            {client.type==='prospect' ? <button style={{ ...btn('primary'), width:'100%', marginBottom:10 }} onClick={()=>startSession('discovery')}>🎯 New Sales Deck</button> : (<><button style={{ ...btn('primary'), width:'100%', marginBottom:10 }} onClick={()=>startSession('monthly')}>📅 Monthly Review</button><button style={{ ...btn('ghost'), width:'100%', marginBottom:10 }} onClick={()=>startSession('quarterly')}>📊 Quarterly Review</button><button style={{ ...btn('ghost'), width:'100%' }} onClick={()=>startSession('annual')}>🏆 Annual Recap</button></>)}
            <div style={{ marginTop:14, padding:12, background:C.surface, borderRadius:8, fontSize:12, color:C.faint, lineHeight:1.6 }}>Each session saves to deck history so you can track progress and learn from past builds.</div>
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
      <div style={{ maxWidth:640 }}>
        <StepBar />
        <div style={st.card}>
          <div style={{ fontSize:16, fontWeight:700, marginBottom:6 }}>Upload Discovery Call Transcript</div>
          <div style={{ fontSize:13, color:C.muted, marginBottom:18 }}>Upload a .txt or .docx transcript. AI will extract pain points, goals, gaps, and key signals.</div>
          <div style={{ border:`2px dashed ${C.border}`, borderRadius:10, padding:24, textAlign:'center', cursor:'pointer', marginBottom:14 }} onClick={()=>transcriptRef.current?.click()}>
            <div style={{ fontSize:32, marginBottom:8 }}>{fileName?'📄':'⬆️'}</div>
            <div style={{ fontSize:14, fontWeight:600, color:fileName?C.success:C.text }}>{fileName||'Click to upload transcript'}</div>
            <div style={{ fontSize:12, color:C.faint, marginTop:4 }}>.txt or .docx supported</div>
            <input ref={transcriptRef} type="file" accept=".txt,.doc,.docx" style={{ display:'none' }} onChange={handleFile} />
          </div>
          <textarea style={{ ...st.textarea, minHeight:120 }} placeholder="Or paste transcript text directly…" value={transcript} onChange={e=>setTranscript(e.target.value)} />
          {transcript && <div style={{ display:'flex', gap:8, padding:'10px 14px', background:'#0F2A1D', border:'1px solid #059669', borderRadius:8, alignItems:'center', margin:'12px 0' }}><span>✅</span><span style={{ fontSize:13, color:C.success }}>{transcript.split(/\s+/).length.toLocaleString()} words ready</span></div>}

          {/* Reference Deck Upload */}
          <div style={{ marginTop:20, padding:16, background:C.surface, borderRadius:10, border:`1px solid ${C.border}` }}>
            <div style={{ fontSize:13, fontWeight:700, marginBottom:4 }}>📎 Upload a Reference Deck <span style={{ fontSize:12, fontWeight:400, color:C.faint }}>(optional)</span></div>
            <div style={{ fontSize:12, color:C.muted, marginBottom:12 }}>Upload a past BTR deck (PDF) to show the AI what quality and narrative depth to match. Same as using a past deck in a Claude Project.</div>
            {refDeckName ? (
              <div style={{ display:'flex', alignItems:'center', justifyContent:'space-between', padding:'10px 14px', background:C.card, borderRadius:8, border:`1px solid ${C.success}` }}>
                <div style={{ display:'flex', gap:8, alignItems:'center' }}><span>📋</span><span style={{ fontSize:13, color:C.success, fontWeight:600 }}>{refDeckName}</span></div>
                <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted, fontSize:11 }} onClick={()=>{ setRefDeck(null); setRefDeckName(''); }}>Remove</button>
              </div>
            ) : (
              <button style={{ ...btn('ghost'), width:'100%' }} onClick={()=>refDeckRef.current?.click()}>⬆ Upload Reference Deck (PDF)</button>
            )}
            <input ref={refDeckRef} type="file" accept=".pdf" style={{ display:'none' }} onChange={handleRefDeck} />
          </div>

          <button style={{ ...btn('primary'), opacity:(!transcript||analyzing)?0.5:1, marginTop:16 }} onClick={doAnalyze} disabled={!transcript||analyzing}>
            {analyzing?'⏳ Analyzing transcript…':'🔍 Analyze with AI'}
          </button>
        </div>
      </div>
    );

    if (step===1 && analysis) return (
      <div>
        <StepBar />
        {analysis.summary && <div style={{ ...st.card, borderLeft:`3px solid ${C.accent}`, marginBottom:16 }}><div style={{ fontSize:11, fontWeight:700, color:C.accent, marginBottom:8, textTransform:'uppercase', letterSpacing:'0.6px' }}>Discovery Brief</div><div style={{ fontSize:14, color:'#E2E8F0', lineHeight:1.7 }}>{analysis.summary}</div></div>}
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
            <div style={{ fontSize:13, color:C.muted, marginTop:4 }}>All optional. Each report analyzed against BTR's diagnostic framework. No numbers fabricated.</div>
          </div>
          <div style={{ display:'flex', gap:8, alignItems:'center', flexShrink:0 }}>
            {uploadedCount>0 && <span style={{ fontSize:12, color:C.success, fontWeight:600 }}>✓ {uploadedCount} uploaded</span>}
            <button style={{ ...btn('ghost'), fontSize:12 }} onClick={doGenerate}>Skip → Build Deck</button>
          </div>
        </div>
        <input ref={reportFileRef} type="file" accept=".csv,.txt,.tsv" style={{ display:'none' }} onChange={handleReportFile} />
        {Object.entries(catGroups).map(([cat,rpts])=>(
          <div key={cat} style={{ marginBottom:20 }}>
            <div style={{ display:'flex', alignItems:'center', gap:8, marginBottom:10 }}>
              <span style={{ width:10, height:10, borderRadius:'50%', background:CAT_COLORS[cat]||C.muted, display:'inline-block' }}></span>
              <span style={{ fontSize:13, fontWeight:700, color:C.muted, textTransform:'uppercase', letterSpacing:'0.6px' }}>{cat}</span>
            </div>
            <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill,minmax(240px,1fr))', gap:10 }}>
              {rpts.map(r=>{ const up=uploadedReports[r.id]; return (
                <div key={r.id} style={{ background:C.card, border:`1px solid ${up?C.success:C.border}`, borderRadius:10, padding:'12px 14px', display:'flex', flexDirection:'column', gap:8 }}>
                  <div style={{ fontSize:13, fontWeight:600, color:up?C.success:C.text }}>{up&&'✓ '}{r.label}</div>
                  {up ? <div><div style={{ fontSize:11, color:C.faint }}>{up.fileName} · {up.rowCount.toLocaleString()} rows</div><button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted, marginTop:6, padding:'3px 10px' }} onClick={()=>removeReport(r.id)}>Remove</button></div>
                      : <button style={{ ...btn('sm'), background:'transparent', border:`1px solid ${C.border}`, color:C.muted }} onClick={()=>triggerReportUpload(r.id)}>⬆ Upload CSV</button>}
                </div>
              ); })}
            </div>
          </div>
        ))}
        {uploadedCount>0 && !reportAnalysis && (
          <div style={{ ...st.card, borderTop:`3px solid ${C.accent}`, marginTop:8 }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
              <div><div style={{ fontSize:15, fontWeight:700 }}>Ready to Analyze</div><div style={{ fontSize:13, color:C.muted, marginTop:3 }}>{uploadedCount} report{uploadedCount>1?'s':''} — BTR diagnostic framework will be applied to each.</div></div>
              <button style={{ ...btn('primary'), opacity:analyzingReports?0.5:1, flexShrink:0, marginLeft:16 }} onClick={doAnalyzeReports} disabled={analyzingReports}>{analyzingReports?'⏳ Analyzing…':'🔍 Run Diagnostic Analysis'}</button>
            </div>
          </div>
        )}
        {reportAnalysis && (
          <div style={{ marginTop:20 }}>
            {(reportAnalysis.strategicRecommendations||[]).length>0 && (
              <div style={{ ...st.card, borderTop:`3px solid ${C.accent}`, marginBottom:16 }}>
                <div style={{ fontSize:13, fontWeight:700, color:C.accent, marginBottom:4, textTransform:'uppercase', letterSpacing:'0.6px' }}>🎯 Strategic Recommendations</div>
                <div style={{ fontSize:12, color:C.muted, marginBottom:16 }}>Problems identified from your data, mapped to BTR's playbook. These will drive the deck narrative.</div>
                {(reportAnalysis.strategicRecommendations||[]).map((rec,i)=>(
                  <div key={i} style={{ marginBottom:14, padding:14, background:C.surface, borderRadius:10, borderLeft:`3px solid ${rec.priority==='high'?C.error:rec.priority==='medium'?C.warn:C.muted}` }}>
                    <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:8 }}>
                      <div style={{ fontSize:14, fontWeight:700, color:C.text }}>{rec.problem}</div>
                      <span style={{ fontSize:11, padding:'2px 8px', borderRadius:20, background:rec.priority==='high'?'#EF444420':rec.priority==='medium'?'#F59E0B20':'#37415120', color:rec.priority==='high'?C.error:rec.priority==='medium'?C.warn:C.muted, fontWeight:600, flexShrink:0, marginLeft:8 }}>{rec.priority}</span>
                    </div>
                    <div style={{ fontSize:12, color:'#94A3B8', marginBottom:8, paddingLeft:10, borderLeft:'1px solid #374151' }}><span style={{ color:C.warn, fontWeight:600 }}>Evidence: </span>{rec.evidence}</div>
                    <div style={{ fontSize:12, color:'#CBD5E1', paddingLeft:10, borderLeft:'1px solid #374151' }}><span style={{ color:C.success, fontWeight:600 }}>BTR Strategy: </span>{rec.btrStrategy}</div>
                  </div>
                ))}
              </div>
            )}
            <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:14, marginBottom:20 }}>
              {[['keyMetrics','Key Metrics','📊',C.blue],['strengthsFound','Strengths Found','✅',C.success],['gapsFound','Gaps Found','⚠️',C.warn],['opportunities','Opportunities','⚡',C.accent],['competitiveInsights','Competitive Insights','🔍','#A78BFA'],['dataLimitations','Data Limitations','🔒',C.faint]].map(([k,l,ic,col])=>(
                reportAnalysis[k]?.length>0 && (
                  <div key={k} style={{ ...st.card, borderTop:`3px solid ${col}` }}>
                    <div style={{ fontSize:11, fontWeight:700, color:col, marginBottom:10, textTransform:'uppercase', letterSpacing:'0.5px' }}>{ic} {l}</div>
                    {(reportAnalysis[k]||[]).map((item,i)=>{
                      const isMetric=k==='keyMetrics';
                      return isMetric
                        ? <div key={i} style={{ marginBottom:10, padding:'8px 10px', background:C.surface, borderRadius:6 }}><div style={{ fontSize:13, fontWeight:700, color:C.text }}>{item.metric}: <span style={{ color:col }}>{item.value}</span></div><div style={{ fontSize:11, color:C.faint, marginTop:2 }}>{item.source} · {item.context}</div></div>
                        : <div key={i} style={{ fontSize:13, color:'#CBD5E1', marginBottom:7, paddingLeft:10, borderLeft:`2px solid ${col}30`, lineHeight:1.5 }}>{item}</div>;
                    })}
                  </div>
                )
              ))}
            </div>
            <div style={{ ...st.card, borderTop:`3px solid ${C.blue}` }}>
              <div style={{ fontSize:14, fontWeight:700, marginBottom:4 }}>💬 Ask Questions About the Data</div>
              <div style={{ fontSize:12, color:C.muted, marginBottom:12 }}>Answers based strictly on uploaded data. No numbers fabricated.</div>
              <div style={{ background:C.surface, borderRadius:10, padding:16, minHeight:160, maxHeight:280, overflowY:'auto', marginBottom:12, display:'flex', flexDirection:'column', gap:10 }}>
                {chatHistory.map((m,i)=>(
                  <div key={i} style={{ display:'flex', justifyContent:m.role==='user'?'flex-end':'flex-start' }}>
                    <div style={{ maxWidth:'82%', padding:'10px 14px', borderRadius:10, fontSize:13, lineHeight:1.6, whiteSpace:'pre-wrap', background:m.role==='user'?C.accent:C.card, color:'#fff', borderBottomRightRadius:m.role==='user'?2:10, borderBottomLeftRadius:m.role==='user'?10:2 }}>{m.content}</div>
                  </div>
                ))}
                {chatLoading && <div style={{ display:'flex', justifyContent:'flex-start' }}><div style={{ padding:'10px 14px', background:C.card, borderRadius:10, fontSize:13, color:C.muted }}>Analyzing…</div></div>}
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
              <button style={{ ...btn('primary'), opacity:generating?0.5:1 }} onClick={doGenerate} disabled={generating}>{generating?'⏳ Building Narrative Audit…':'✨ Generate Deck from Analysis'}</button>
            </div>
          </div>
        )}
        {uploadedCount===0 && <div style={{ display:'flex', gap:10, marginTop:8 }}><button style={{ ...btn('primary'), opacity:generating?0.5:1 }} onClick={doGenerate} disabled={generating}>{generating?'⏳ Building…':'✨ Build Deck (Discovery Only)'}</button></div>}
      </div>
    );

    if (step===3 && slides.length>0) return (
      <div>
        <StepBar />
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:18 }}>
          <div style={{ fontSize:13, color:C.muted }}>{slides.length} slides · click any slide to edit</div>
          <div style={{ display:'flex', gap:10 }}>
            <button style={btn('ghost')} onClick={()=>setStep(2)}>← Reports</button>
            <button style={{ ...btn('primary'), opacity:savingSession?0.5:1 }} onClick={saveAndAdvance} disabled={savingSession}>{savingSession?'Saving…':'Save & Export →'}</button>
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
                <><div style={{ fontSize:14, fontWeight:700, marginBottom:5 }}>{sl.title}</div><div style={{ fontSize:12, color:C.accent, fontStyle:'italic', marginBottom:8 }}>{sl.headline}</div>{(sl.bullets||[]).slice(0,3).map((b,bi)=><div key={bi} style={{ fontSize:12, color:C.muted, marginBottom:4, paddingLeft:12, position:'relative' }}><span style={{ position:'absolute', left:0, color:C.accent }}>•</span>{b}</div>)}{(sl.bullets||[]).length>3&&<div style={{ fontSize:11, color:C.faint }}>+{sl.bullets.length-3} more</div>}</>
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
            <div style={{ fontSize:20, fontWeight:800, marginBottom:6 }}>Deck Saved</div>
            <div style={{ fontSize:14, color:C.muted, marginBottom:22 }}>{slides.length} slides · saved to {client?.name}'s history</div>
            <button style={{ ...btn('primary'), width:'100%', padding:'13px 20px', fontSize:15, marginBottom:10 }} onClick={exportPPTX} disabled={exporting}>{exporting?'⏳ Exporting…':'⬇️ Download PowerPoint (.pptx)'}</button>
          </div>
          <div style={{ borderTop:`1px solid ${C.border}`, paddingTop:16 }}>
            <div style={{ fontSize:13, fontWeight:700, marginBottom:10 }}>Slide List</div>
            {slides.map((sl,i)=><div key={sl.id} style={{ display:'flex', gap:10, alignItems:'center', padding:'6px 0', borderBottom:'1px solid #1F2937' }}><span style={{ fontSize:11, color:C.faint, width:18 }}>{i+1}</span><span style={{ fontSize:14 }}>{slideIcons[sl.type]||'📄'}</span><span style={{ fontSize:13, color:'#CBD5E1' }}>{sl.title}</span></div>)}
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
  const meta = { clients:{title:'Client Manager',sub:'All prospects and active clients'}, 'new-client':{title:'New Client',sub:'Create a client profile'}, profile:{title:client?.name||'',sub:`${client?.type==='prospect'?'Prospect':'Active Client'} · ${client?.categories||''}`}, session:{title:'Deck Builder',sub:`${client?.name||''} · ${sessionType==='discovery'?'Sales Deck':sessionType.charAt(0).toUpperCase()+sessionType.slice(1)+' Review'}`} };
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
        <div style={{ padding:'14px 20px', borderTop:'1px solid #1F2937', fontSize:11, color:'#374151' }}>v4.0 · Narrative Engine</div>
      </div>
      <div style={st.main}>
        <div style={st.header}><div style={{ fontSize:19, fontWeight:700 }}>{title}</div>{sub&&<div style={{ fontSize:13, color:C.faint, marginTop:3 }}>{sub}</div>}</div>
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
