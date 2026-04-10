# PBV Variance Narrator v1.2 — Agent Resume

## Agent Profile
- Name: PBV Variance Narrator
- Version: 1.2
- Developer: Bhargav Venkatesh (PBV Finance)
- Certifications: CMA, FMVA, BIDA

## Architecture
- 3-Layer: Data Ingestion → Accounting Intelligence → FP&A Engine
- 3 Agents: Calculator (Python) → Diagnostician (Gemma 4) → Memo Writer (Gemma 4)
- Decision Engine: Rule-based Category 1-4 (Python, not AI)

## Anti-Hallucination
- Hardcoded country rules (UAE, India, KSA, Qatar, UK)
- Pre-built quick wins library (15+ verified actions with SAP T-codes)
- Confidence scoring (CALCULATED / PRE-BUILT / HYPOTHESIS)
- 10-point self-validation
- GL code intelligence for auto-mapping
- Mapping memory for repeat accuracy

## Performance
- Agent 1: <1 second (any TB size)
- Agent 2: 5-10 min (Gemma 4, 8GB RAM)
- Agent 3: 5-10 min (Gemma 4, 8GB RAM)
- Accuracy: Agent 1 = 100%, Overall with review = 95%+

## Capabilities
- ANY trial balance format (auto-detects structure)
- GL code mapping (SAP ranges)
- P&L vs Balance Sheet separation
- Multi-entity, multi-period support
- 6 variance tables + EBITDA bridge
- CFO board memo (10 sections)
- Country-specific compliance flags
- Excel report download

## Deployment
- Local: Ollama + Gemma 4 (zero cloud)
- Cloud: Streamlit (Agent 1 only)
- Hybrid: local AI + cloud interface

## Contact
- LinkedIn: linkedin.com/in/bhargavvenkatesh
- Demo: [Streamlit URL]
- Video: [Loom URL]
