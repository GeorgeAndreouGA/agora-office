// Formula engine worker: recalculation runs off the main thread so typing stays smooth.
import { UniverSheetsCoreWorkerPreset } from '@univerjs/preset-sheets-core/worker';
import { createUniver, LocaleType } from '@univerjs/presets';

createUniver({
  locale: LocaleType.EN_US,
  presets: [UniverSheetsCoreWorkerPreset()],
});
