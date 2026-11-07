import type { ComponentType } from 'react';
import type { DocKind } from './config';
import type { EditorProps } from './types';

type EditorModule = { default: ComponentType<EditorProps> };

/** Each editor is its own download, fetched only when chosen (or hovered on the hub). */
export const editorLoaders: Record<DocKind, () => Promise<EditorModule>> = {
  docx: () => import('../papyros/Papyros'),
  xlsx: () => import('../abakion/Abakion'),
  pptx: () => import('../pinax/Pinax'),
};

export function preloadEditor(kind: DocKind) {
  void editorLoaders[kind]().catch(() => undefined);
}
