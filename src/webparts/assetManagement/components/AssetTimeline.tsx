import * as React from 'react';
import { useState } from 'react';
import styles from './AssetManagement.module.scss';
import { selectInitialTimelineEvents, type IAssetTimelineEvent } from './assetTimelineService';

interface IProps { events: IAssetTimelineEvent[]; }
const typeLabel: Record<IAssetTimelineEvent['type'], string> = {
  purchased: 'Purchased', movement: 'Moved / allocated', failure: 'Failure', resolved: 'Resolved / closed',
  'support-expiring': 'Support expiring', 'support-expired': 'Support expired', 'support-valid': 'Support valid'
};
const toneClasses: Record<IAssetTimelineEvent['type'], string> = {
  purchased: styles.timelineTonePurchased,
  movement: styles.timelineToneMovement,
  failure: styles.timelineToneFailure,
  resolved: styles.timelineToneResolved,
  'support-expiring': styles.timelineToneSupportExpiring,
  'support-expired': styles.timelineToneSupportExpired,
  'support-valid': styles.timelineToneSupportValid
};
const AssetTimeline: React.FC<IProps> = ({ events }) => {
  const [showAll, setShowAll] = useState(false);
  const visible = showAll ? events : selectInitialTimelineEvents(events);
  if (!events.length) return <p className={styles.timelineEmpty}>No history yet for this asset.</p>;
  return <section className={styles.timelineSection} aria-label="Asset timeline">
    <ul className={styles.timelineLegend} aria-label="Timeline event legend">
      {(Object.keys(typeLabel) as Array<keyof typeof typeLabel>).map(type => <li key={type}><i className={`${styles.timelineSwatch} ${toneClasses[type]}`}/>{typeLabel[type]}</li>)}
    </ul>
    <div className={styles.timelineScroll} tabIndex={0} aria-label="Timeline, scroll sideways for more">
      <ol className={styles.timelineRail}>
        {visible.map(event => <li key={event.id} className={`${styles.timelineEvent} ${toneClasses[event.type]} ${event.position === 'above' ? styles.timelineAbove : styles.timelineBelow}`}>
          <article className={`${styles.timelineCard} ${toneClasses[event.type]}`} aria-label={`${event.date.toLocaleDateString(undefined, { month: 'long', year: 'numeric' })}, ${event.title}, ${event.type === 'failure' ? 'reported on' : event.type === 'resolved' ? 'resolved on' : 'on'} ${event.date.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' })}`}>
            <header><strong>{event.date.getFullYear()}</strong><span>{event.date.toLocaleDateString(undefined, { month: 'short' }).toUpperCase()}</span></header>
            <div><h4>{event.title}</h4><p>{event.description}</p><time dateTime={event.date.toISOString()}>{event.type === 'failure' ? 'Reported on ' : event.type === 'resolved' ? 'Resolved on ' : 'On '}{event.date.toLocaleDateString(undefined, { day: '2-digit', month: 'short', year: 'numeric' })}</time></div>
          </article>
          <span className={styles.timelineMarker} aria-hidden="true"><i/></span>
        </li>)}
      </ol>
    </div>
    {events.length > 8 && <button type="button" className={styles.timelineToggle} aria-expanded={showAll} onClick={() => setShowAll(value => !value)}>{showAll ? 'Show fewer' : `Show all (${events.length})`}</button>}
    <p className={styles.timelineCaption}>Generated from the Asset Tracker and service history | Renewal mail is sent to the Asset Owner and Asset Custodian after the Expiry Date Of Support.</p>
  </section>;
};

export default AssetTimeline;
