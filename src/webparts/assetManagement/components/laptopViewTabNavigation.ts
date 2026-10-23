export type LaptopViewTabKey = 'details' | 'timeline' | 'attachments' | 'service';
export const laptopViewTabKeys: LaptopViewTabKey[] = ['details', 'timeline', 'attachments', 'service'];

export const getNextLaptopViewTab = (current: LaptopViewTabKey, key: string): LaptopViewTabKey | undefined => {
  const index = laptopViewTabKeys.indexOf(current);
  if (key === 'Home') return laptopViewTabKeys[0];
  if (key === 'End') return laptopViewTabKeys[laptopViewTabKeys.length - 1];
  if (key === 'ArrowRight') return laptopViewTabKeys[(index + 1) % laptopViewTabKeys.length];
  if (key === 'ArrowLeft') return laptopViewTabKeys[(index - 1 + laptopViewTabKeys.length) % laptopViewTabKeys.length];
  return undefined;
};
