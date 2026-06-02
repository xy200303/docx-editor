import { describe, test, expect } from 'bun:test';
import type { EditorBridge } from '../bridge';
import { executeToolCall } from '../tools';

function makeBridge(overrides: Partial<EditorBridge> = {}): EditorBridge {
  return {
    setContentControlValue: () => true,
    ...overrides,
  } as EditorBridge;
}

describe('set_content_control_value', () => {
  test('sets typed SDT values by stable anchors', () => {
    type SetValueOptions = Parameters<EditorBridge['setContentControlValue']>[0];
    const cases: Array<{ input: Record<string, unknown>; expected: SetValueOptions }> = [
      {
        input: { tag: 'status', kind: 'dropdown', value: 'Approved' },
        expected: {
          tag: 'status',
          value: { kind: 'dropdown', value: 'Approved' },
          force: undefined,
        },
      },
      {
        input: { alias: 'Agree', kind: 'checkbox', checked: true, force: true },
        expected: {
          alias: 'Agree',
          value: { kind: 'checkbox', checked: true },
          force: true,
        },
      },
      {
        input: { id: 12, kind: 'date', date: '2026-06-01' },
        expected: {
          id: 12,
          value: { kind: 'date', date: '2026-06-01' },
          force: undefined,
        },
      },
    ];

    for (const { input, expected } of cases) {
      let captured: Parameters<EditorBridge['setContentControlValue']>[0] | undefined;
      const bridge = makeBridge({
        setContentControlValue: (opts) => {
          captured = opts;
          return true;
        },
      });
      const result = executeToolCall('set_content_control_value', input, bridge);
      expect(result.success).toBe(true);
      expect(captured).toEqual(expected);
    }
  });

  test('requires an SDT anchor', () => {
    const result = executeToolCall(
      'set_content_control_value',
      { kind: 'checkbox', checked: true },
      makeBridge()
    );
    expect(result.success).toBe(false);
    expect(result.error).toContain('tag');
  });

  test('validates kind-specific inputs', () => {
    const checkbox = executeToolCall(
      'set_content_control_value',
      { tag: 'agree', kind: 'checkbox', checked: 'yes' },
      makeBridge()
    );
    expect(checkbox.success).toBe(false);
    expect(checkbox.error).toContain('checked');

    const date = executeToolCall(
      'set_content_control_value',
      { tag: 'effective', kind: 'date', date: '06/01/2026' },
      makeBridge()
    );
    expect(date.success).toBe(false);
    expect(date.error).toContain('yyyy-mm-dd');
  });
});
