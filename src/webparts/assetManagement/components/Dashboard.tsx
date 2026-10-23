import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import { Icon } from '@fluentui/react';
import styles from './AssetManagement.module.scss';
import { calculateDashboardSummary, expiryPillTone, expiringAssets, loadDashboardData, type IDashboardData, type IDashboardAsset } from './dashboardDataService';
import type { SPHttpClient } from '@microsoft/sp-http';

interface IProps { client: SPHttpClient; onNavigate: (registerKey: string, status?: string, itemIds?: number[]) => void; onRegisterCounts: (counts: Record<string, number>) => void; }
const statuses = ['Allocated', 'In Stock', 'In Maintenance', 'Retired'] as const;
const WidgetError: React.FC<{ items: IDashboardData['registers']; retry: () => void }> = ({ items, retry }) => items.length ? <div className={styles.dashboardError} role="alert">{items.map(item => <p key={item.config.key}>{item.config.name}: {item.error}</p>)}<button type="button" onClick={retry}>Retry</button></div> : null;
const Dashboard: React.FC<IProps> = ({ client, onNavigate, onRegisterCounts }) => {
  const [data, setData] = useState<IDashboardData | null>(null);
  const [loading, setLoading] = useState(true);
  const [updatedAt, setUpdatedAt] = useState<Date | null>(null);
  const [windowDays, setWindowDays] = useState(30);
  const [expandedWhy, setExpandedWhy] = useState(false);
  const refresh = async (force = false): Promise<void> => { setLoading(true); const next = await loadDashboardData(client, force); setData(next); setUpdatedAt(next.updatedAt); setLoading(false); };
  useEffect(() => { void refresh(); }, [client]);
  useEffect(() => { if (data) onRegisterCounts(data.registers.reduce<Record<string, number>>((counts, item) => { counts[item.config.key] = item.assets.length; return counts; }, {})); }, [data, onRegisterCounts]);
  const summary = useMemo(() => data ? calculateDashboardSummary(data, windowDays) : undefined, [data, windowDays]);
  const failed = (predicate: (config: IDashboardData['registers'][number]['config']) => boolean): IDashboardData['registers'] => data?.registers.filter(item => !!item.error && predicate(item.config)) || [];
  const statusesErrors = failed(config => !!config.status);
  const expiryErrors = failed(config => !!config.expiry?.length);
  const stockErrors = failed(config => !!config.status || !!config.totalCount);
  const dataErrors = failed(() => true);
  const expirableRecords = data ? data.registers.filter(item => !item.error && item.hasExpiry).reduce<IDashboardAsset[]>((all, item) => all.concat(item.assets), []) : [];
  const visibleExpiry = data ? expiringAssets(expirableRecords, windowDays) : [];
  const top = summary?.byRegister.slice(0, 6) || [];
  const others = summary?.byRegister.slice(6) || [];
  const perRegister = summary?.byRegister || [];
  const maxCount = Math.max(1, ...top.map(item => item.assets.length));
  const chartCounts = statuses.map(status => summary?.statuses[status] || 0);
  const chartTotal = chartCounts.reduce((sum, count) => sum + count, 0);
  const circumference = 2 * Math.PI * 48;
  let accumulated = 0;
  const arcs = statuses.map((status, index) => { const length = chartTotal ? chartCounts[index] / chartTotal * circumference : 0; const arc = { status, count: chartCounts[index], length, offset: accumulated }; accumulated += length; return arc; });
  const widget = (busy: boolean, children: React.ReactNode, errors: IDashboardData['registers'] = []): React.ReactElement => <section className={styles.dashboardWidget}>
    {busy && !data ? <div className={styles.dashboardSkeleton} role="status" aria-label="Loading dashboard data"/> : <>{children}<WidgetError items={errors} retry={() => void refresh(true)}/></>}
  </section>;
  const formatDate = (value: string): string => {
    const match = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    return match ? new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3])).toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' }) : 'Date missing';
  };
  return <div className={styles.dashboardPage}>
    <header className={styles.dashboardHeader}><div><h1>Dashboard</h1><p>Live overview of assets, warranties and stock across your registers.</p></div><div><button type="button" className={styles.secondaryButton} onClick={() => void refresh(true)} disabled={loading}><Icon iconName="Refresh"/>Refresh</button><small>{updatedAt ? `Updated ${updatedAt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}` : 'Loading live data'}</small></div></header>
    <div className={styles.dashboardKpis}>{widget(loading, <div className={styles.dashboardKpiGrid}>
      {[
        { name: 'Total assets', value: summary?.total || 0, className: '' },
        { name: 'Allocated', value: summary?.statuses.Allocated || 0, className: styles.dashboardAllocated },
        { name: 'In stock', value: summary?.statuses['In Stock'] || 0, className: styles.dashboardInStock },
        { name: 'Warranty expiring (next 30 days)', value: expiringAssets(expirableRecords, 30).length, className: styles.metricMaintenance },
        { name: 'Warranty expired', value: summary?.expired.length || 0, className: styles.metricRetired }
      ].map(card => <button key={card.name} type="button" aria-disabled="true" tabIndex={-1} className={styles.dashboardKpi}><span>{card.name}</span><strong className={card.className}>{card.value}</strong></button>)}
    </div>, dataErrors)}</div>
    {summary && (summary.noExpiryDate > 0 || summary.duplicateIds.length > 0 || data?.registers.some(item => item.warning)) && <aside className={styles.dashboardWarnings}><div>{summary.noExpiryDate > 0 && <p>{summary.noExpiryDate} records have no expiry date.</p>}{summary.duplicateIds.length > 0 && <p>{summary.duplicateIds.length} Asset IDs appear on more than one record (counted as separate records).</p>}{data?.registers.filter(item => item.warning).map(item => <p key={item.config.key}>{item.config.name}: {item.warning}</p>)}</div>{summary.duplicateIds.length > 0 && <><button type="button" aria-expanded={expandedWhy} onClick={() => setExpandedWhy(value => !value)}>Why?</button>{expandedWhy && <ul>{summary.duplicateIds.map(item => <li key={item.id}>{item.id}: {item.count} records</li>)}</ul>}</>}</aside>}
    <div className={styles.dashboardCharts}>
      {widget(loading, <><h2>Asset status</h2><div className={styles.dashboardDonutLayout}><svg viewBox="0 0 120 120" role="img" aria-label={`Asset status counts: ${statuses.map((status, index) => `${status} ${chartCounts[index]}`).join(', ')}`}><circle cx="60" cy="60" r="48" fill="none" stroke="currentColor" strokeWidth="14" opacity=".12"/>{arcs.map(arc => <circle key={arc.status} cx="60" cy="60" r="48" fill="none" className={styles[`donut${arc.status.replace(/\s/g, '')}` as keyof typeof styles]} strokeWidth="14" strokeDasharray={`${arc.length} ${circumference - arc.length}`} strokeDashoffset={-arc.offset} transform="rotate(-90 60 60)"/>)}<text x="60" y="57" textAnchor="middle" className={styles.donutNumber}>{chartTotal}</text><text x="60" y="71" textAnchor="middle" className={styles.donutLabel}>assets</text></svg><ul className={styles.dashboardLegend}>{statuses.map((status, index) => <li key={status}><button type="button" aria-disabled="true" tabIndex={-1}><i className={styles[`donut${status.replace(/\s/g, '')}` as keyof typeof styles]}/><span>{status}</span><strong>{chartCounts[index]}</strong></button></li>)}{(summary?.statuses.Unspecified || 0) > 0 && <li><button type="button" aria-disabled="true" tabIndex={-1}><i className={styles.donutOther}/><span>Unspecified</span><strong>{summary?.statuses.Unspecified}</strong></button></li>}{(summary?.statuses.Other || 0) > 0 && <li><span>Other status</span><strong>{summary?.statuses.Other}</strong></li>}</ul></div><small>Unspecified and other statuses are excluded from the four status percentages.</small></>, statusesErrors)}
      {widget(loading, <><h2>Assets by register</h2><ul className={styles.dashboardRegisterBars}>{top.map(item => { const counts = ['Allocated', 'In Stock', 'In Maintenance', 'Retired'].map(key => item.statuses[key] || 0); const unknown = Math.max(0, item.assets.length - counts.reduce((sum, value) => sum + value, 0)); const colors = [styles.donutAllocated, styles.donutInStock, styles.donutInMaintenance, styles.donutRetired, styles.donutOther]; const values = [...counts, unknown]; return <li key={item.config.key}><button type="button" aria-disabled="true" tabIndex={-1}><span>{item.config.name}</span><div className={styles.dashboardBarTrack} role="img" aria-label={`${item.assets.length} assets`}>{values.map((value, index) => value > 0 && <i key={index} className={colors[index]} style={{ width: `${value / maxCount * 100}%` }}/>)}</div><strong>{item.assets.length}</strong></button></li>; })}{others.length > 0 && <li><button type="button" aria-disabled="true" tabIndex={-1}><span>Other {others.length} registers</span><div className={styles.dashboardBarTrack} role="img" aria-label={`${others.reduce((sum, item) => sum + item.assets.length, 0)} assets`}/><strong>{others.reduce((sum, item) => sum + item.assets.length, 0)}</strong></button></li>}</ul><div className={styles.dashboardBarLegend}>{statuses.map(status => `${status}: ${perRegister.reduce((sum, item) => sum + (item.statuses[status] || 0), 0)}`).join(' · ')}</div></>, dataErrors)}
    </div>
    <div className={styles.dashboardTables}>
      {widget(loading, <><div className={styles.dashboardWidgetHeader}><h2>Expiring warranties</h2><div className={styles.dashboardWindows} role="group" aria-label="Warranty expiry period">{[7, 15, 30].map(days => <button type="button" key={days} aria-pressed={windowDays === days} onClick={() => setWindowDays(days)}>{days} days</button>)}</div></div><div className={styles.dashboardTableScroll}><table><thead><tr><th>Asset</th><th>Register</th><th>Type</th><th>Assigned to</th><th>Support ends</th><th>Days left</th></tr></thead><tbody>{visibleExpiry.slice(0, 10).map(asset => <tr key={`${asset.registerKey}-${asset.itemId}`}><td><span>{asset.name}<small>{asset.assetId || 'No Asset ID'}</small></span></td><td>{asset.registerName}</td><td>{asset.expiryType}</td><td>{asset.assignedTo || asset.custodian || asset.location || 'Unassigned'}</td><td>{formatDate(asset.expiry)}</td><td><span className={`${styles.dashboardPill} ${styles[`expiry${expiryPillTone(asset.daysLeft)}`]}`}>{asset.daysLeft} days</span></td></tr>)}{!visibleExpiry.length && !loading && !expiryErrors.length && <tr><td colSpan={6} className={styles.dashboardEmpty}>No warranties expire in this period.</td></tr>}</tbody></table></div></>, expiryErrors)}
      {widget(loading, <><h2>Stock details</h2><div className={styles.dashboardTableScroll}><table><thead><tr><th>Register</th><th>In stock</th><th>Allocated</th><th>Level</th></tr></thead><tbody>{(summary?.stocks || []).map(stock => { return <tr key={stock.config.key}><td><span>{stock.config.name}{stock.config.totalCount && <small>{stock.total} total · View in stock</small>}</span></td><td>{stock.inStock}</td><td>{stock.allocated}</td><td><span className={`${styles.dashboardPill} ${styles[`stock${stock.level.replace(/\s/g, '')}` as keyof typeof styles]}`}>{stock.level}</span></td></tr>; })}{!summary?.stocks.length && !loading && <tr><td colSpan={4} className={styles.dashboardEmpty}>No stock data available.</td></tr>}</tbody></table></div></>, stockErrors)}
    </div>
  </div>;
};

export default Dashboard;
