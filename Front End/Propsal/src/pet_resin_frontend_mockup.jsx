import React from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { ArrowRight, TrendingUp, Scale, ShieldCheck, Calculator, BarChart3, Sparkles, Target } from 'lucide-react';

const kpiData = [
  { label: 'Total PET Spend', value: '$312.0M', delta: '+4.1% YoY' },
  { label: 'Market TLC Gap', value: '−$18.4/MT', delta: 'vs supplier price' },
  { label: 'Savings Opportunity', value: '$6.8M', delta: 'annualized' },
  { label: 'Watchlist Suppliers', value: '7', delta: 'over market' },
];

const marketRows = [
  ['Brazil', '$1,218', '$1,280', '-$62'],
  ['Argentina', '$1,117', '$921', '+$196'],
  ['Colombia', '$815', '$954', '-$139'],
  ['Peru', '$875', '$1,024', '-$149'],
];

const supplierRows = [
  ['Supplier A', 'Brazil', '$1,295', '$1,218', '+$77'],
  ['Supplier B', 'Argentina', '$948', '$921', '+$27'],
  ['Supplier C', 'China', '$1,006', '$815', '+$191'],
  ['Supplier D', 'USA', '$1,030', '$1,024', '+$6'],
];

const scenarioRows = [
  ['Base case', '$1,218', '$312.0M', '—'],
  ['Resin +10%', '$1,336', '$321.1M', '+$9.1M'],
  ['Freight +20%', '$1,256', '$315.8M', '+$3.8M'],
  ['Brazil duty +5%', '$1,281', '$318.4M', '+$6.4M'],
];

function FlowStep({ icon: Icon, title, subtitle, active = false }) {
  return (
    <div className={`flex items-start gap-3 rounded-2xl border p-4 shadow-sm ${active ? 'border-black bg-black text-white' : 'border-neutral-200 bg-white'}`}>
      <div className={`rounded-xl p-2 ${active ? 'bg-white text-black' : 'bg-neutral-100 text-neutral-700'}`}>
        <Icon size={18} />
      </div>
      <div className="min-w-0">
        <div className="text-sm font-semibold">{title}</div>
        <div className={`text-xs leading-5 ${active ? 'text-neutral-200' : 'text-neutral-500'}`}>{subtitle}</div>
      </div>
    </div>
  );
}

function SimpleTable({ headers, rows }) {
  return (
    <div className="overflow-hidden rounded-2xl border border-neutral-200 bg-white shadow-sm">
      <table className="w-full text-left text-sm">
        <thead className="bg-neutral-50 text-neutral-500">
          <tr>
            {headers.map((h) => (
              <th key={h} className="px-4 py-3 font-medium">{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, idx) => (
            <tr key={idx} className="border-t border-neutral-100">
              {row.map((cell, j) => (
                <td key={j} className={`px-4 py-3 ${j === row.length - 1 ? 'font-semibold' : ''}`}>
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

export default function PETResinFrontendMockup() {
  return (
    <div className="min-h-screen bg-neutral-50 p-6 text-neutral-900">
      <div className="mx-auto max-w-7xl space-y-6">
        <div className="rounded-3xl bg-gradient-to-r from-neutral-900 to-neutral-700 p-8 text-white shadow-xl">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
            <div>
              <div className="mb-2 inline-flex items-center gap-2 rounded-full bg-white/10 px-3 py-1 text-xs font-medium text-white/90">
                <Sparkles size={14} /> PET Resin Procurement Cockpit
              </div>
              <h1 className="text-3xl font-bold tracking-tight lg:text-5xl">Joao’s decision flow</h1>
              <p className="mt-3 max-w-3xl text-sm leading-6 text-white/80 lg:text-base">
                A front-end mockup for monthly market visibility, supplier reconciliation, and annual sourcing strategy support.
              </p>
            </div>
            <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
              {kpiData.map((kpi) => (
                <div key={kpi.label} className="rounded-2xl bg-white/10 p-4 backdrop-blur">
                  <div className="text-xs text-white/70">{kpi.label}</div>
                  <div className="mt-2 text-2xl font-bold">{kpi.value}</div>
                  <div className="mt-1 text-xs text-white/70">{kpi.delta}</div>
                </div>
              ))}
            </div>
          </div>
        </div>

        <Card className="rounded-3xl shadow-sm">
          <CardHeader>
            <CardTitle className="text-xl">End-to-end flow</CardTitle>
          </CardHeader>
          <CardContent>
            <div className="grid gap-3 lg:grid-cols-5">
              <FlowStep icon={TrendingUp} title="1. Market inputs" subtitle="ICIS, Argus, freight, tariffs, duties" />
              <FlowStep icon={Calculator} title="2. TLC engine" subtitle="Calculate landed cost by route and month" active />
              <FlowStep icon={Scale} title="3. Reconciliation" subtitle="Supplier price vs market-implied TLC" />
              <FlowStep icon={BarChart3} title="4. Scenario view" subtitle="What-if shifts in resin, freight, duty" />
              <FlowStep icon={Target} title="5. Strategy" subtitle="Annual sourcing, negotiation, tolling lever" />
            </div>
          </CardContent>
        </Card>

        <div className="grid gap-6 lg:grid-cols-3">
          <Card className="rounded-3xl lg:col-span-2">
            <CardHeader className="flex flex-row items-center justify-between space-y-0">
              <CardTitle>1) Executive overview</CardTitle>
              <Badge variant="secondary">Monthly refresh</Badge>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="grid gap-4 md:grid-cols-4">
                {kpiData.map((kpi) => (
                  <div key={kpi.label} className="rounded-2xl border border-neutral-200 p-4">
                    <div className="text-xs text-neutral-500">{kpi.label}</div>
                    <div className="mt-2 text-2xl font-bold">{kpi.value}</div>
                    <div className="mt-1 text-xs text-neutral-500">{kpi.delta}</div>
                  </div>
                ))}
              </div>
              <div className="rounded-2xl border border-neutral-200 p-4">
                <div className="mb-3 flex items-center justify-between">
                  <div className="font-semibold">Spend vs market gap</div>
                  <div className="text-xs text-neutral-500">Last 6 months</div>
                </div>
                <div className="grid grid-cols-6 gap-2 text-xs">
                  {['Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan'].map((m, i) => (
                    <div key={m} className="space-y-2 text-center">
                      <div className="h-24 rounded-xl bg-neutral-100">
                        <div
                          className="mt-auto rounded-b-xl bg-black"
                          style={{ height: `${35 + i * 8}%`, marginTop: `${100 - (35 + i * 8)}%` }}
                        />
                      </div>
                      <div className="text-neutral-500">{m}</div>
                    </div>
                  ))}
                </div>
                <div className="mt-3 text-xs text-neutral-500">Example: supplier price stayed above market by $12–$28/MT in the last four months.</div>
              </div>
            </CardContent>
          </Card>

          <Card className="rounded-3xl">
            <CardHeader>
              <CardTitle>Alerts</CardTitle>
            </CardHeader>
            <CardContent className="space-y-3 text-sm">
              <div className="rounded-2xl border border-amber-200 bg-amber-50 p-4">
                <div className="font-semibold">3 suppliers above market</div>
                <div className="mt-1 text-neutral-600">Potential negotiation upside: $2.3M</div>
              </div>
              <div className="rounded-2xl border border-blue-200 bg-blue-50 p-4">
                <div className="font-semibold">Brazil route most protected</div>
                <div className="mt-1 text-neutral-600">Import duty + tax stack drives TLC materially higher.</div>
              </div>
              <div className="rounded-2xl border border-emerald-200 bg-emerald-50 p-4">
                <div className="font-semibold">Mexico/US remains a benchmark</div>
                <div className="mt-1 text-neutral-600">Useful as a tolling reference in negotiations.</div>
              </div>
            </CardContent>
          </Card>
        </div>

        <div className="grid gap-6 lg:grid-cols-2">
          <Card className="rounded-3xl">
            <CardHeader>
              <CardTitle>2) Market view</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <SimpleTable
                headers={['Destination', 'FOB', 'Freight', 'TLC delta']}
                rows={marketRows}
              />
              <div className="rounded-2xl bg-neutral-50 p-4 text-sm text-neutral-600">
                Asia is shown as the low-cost benchmark in several lanes, while Brazil carries a heavier tax layer. This card helps Joao read the market monthly.
              </div>
            </CardContent>
          </Card>

          <Card className="rounded-3xl">
            <CardHeader>
              <CardTitle>3) Supplier vs market reconciliation</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <SimpleTable
                headers={['Supplier', 'Origin', 'Supplier price', 'Market TLC', 'Gap']}
                rows={supplierRows}
              />
              <div className="rounded-2xl border border-neutral-200 p-4">
                <div className="mb-2 text-sm font-semibold">Negotiation view</div>
                <div className="grid grid-cols-3 gap-3 text-sm">
                  <div className="rounded-xl bg-neutral-50 p-3">
                    <div className="text-neutral-500">At risk</div>
                    <div className="text-lg font-bold">2 suppliers</div>
                  </div>
                  <div className="rounded-xl bg-neutral-50 p-3">
                    <div className="text-neutral-500">At parity</div>
                    <div className="text-lg font-bold">1 supplier</div>
                  </div>
                  <div className="rounded-xl bg-neutral-50 p-3">
                    <div className="text-neutral-500">Below market</div>
                    <div className="text-lg font-bold">1 supplier</div>
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>
        </div>

        <div className="grid gap-6 lg:grid-cols-2">
          <Card className="rounded-3xl">
            <CardHeader>
              <CardTitle>4) Scenario analytics</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4">
              <div className="flex flex-wrap gap-2">
                <Badge>Resin +10%</Badge>
                <Badge>Freight +20%</Badge>
                <Badge>Brazil duty +5%</Badge>
                <Badge>Asia lead time +15d</Badge>
              </div>
              <SimpleTable
                headers={['Scenario', 'TLC / MT', 'Annual spend', 'Impact']}
                rows={scenarioRows}
              />
            </CardContent>
          </Card>

          <Card className="rounded-3xl">
            <CardHeader>
              <CardTitle>5) Sourcing strategy recommendation</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4 text-sm">
              <div className="rounded-2xl border border-neutral-200 p-4">
                <div className="font-semibold">Mercosur</div>
                <div className="mt-1 text-neutral-600">Prioritize local supply in Brazil; keep a partial Asian benchmark for leverage in Argentina and Uruguay.</div>
              </div>
              <div className="rounded-2xl border border-neutral-200 p-4">
                <div className="font-semibold">Central America</div>
                <div className="mt-1 text-neutral-600">0% duty market — compare Asia vs US/Mexico depending on export requirements.</div>
              </div>
              <div className="rounded-2xl border border-neutral-200 p-4">
                <div className="font-semibold">Andean / Pacific</div>
                <div className="mt-1 text-neutral-600">Asia is cost-efficient; use Mexico as a spot lever for volatility and working capital.</div>
              </div>
              <div className="rounded-2xl bg-black p-4 text-white">
                <div className="text-xs uppercase tracking-wide text-white/60">Recommended next action</div>
                <div className="mt-1 text-lg font-semibold">Review 3 supplier lanes in Q2 and take 1 contract to RFQ.</div>
              </div>
            </CardContent>
          </Card>
        </div>

        <Card className="rounded-3xl border-dashed border-neutral-300 bg-white/70">
          <CardContent className="flex flex-col gap-4 p-6 md:flex-row md:items-center md:justify-between">
            <div>
              <div className="text-sm font-semibold">Suggested user journey</div>
              <div className="mt-1 text-sm text-neutral-500">Open overview → inspect market → reconcile supplier → run scenario → export negotiation pack.</div>
            </div>
            <div className="flex items-center gap-2 text-sm text-neutral-500">
              <span className="rounded-full bg-neutral-100 px-3 py-1">Dashboard</span>
              <ArrowRight size={16} />
              <span className="rounded-full bg-neutral-100 px-3 py-1">Analytics</span>
              <ArrowRight size={16} />
              <span className="rounded-full bg-neutral-100 px-3 py-1">Decision</span>
            </div>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}
