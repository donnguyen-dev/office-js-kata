# Verify contract

Every problem has a `tests/N-{slug}.test.ts` that defines the *automated* half of pass/fail. The `./verify N` script runs vitest, then prints the manual checklist from the problem file for the *manual* half.

## Why the split

Real Word behaviour — event timing, selection state, content control lifecycle, revision tracking — doesn't reliably reproduce in `office-addin-mock`. Attempting to fully automate would produce false negatives (mock behaviour diverges from Word) or false positives (mock says pass, Word doesn't). Neither is useful.

The honest split:

- **Automated** — pure logic, proxy call sequence, argument shape, TypeScript compilation. Mocks are perfectly good at these.
- **Manual** — everything requiring real Word to observe. Confirmed by the learner sideloading into Word Web and following the checklist.

Both must pass. The automated half is machine-gated; the manual half is trust-gated.

## Vitest file shape

```typescript
// tests/N-{slug}.test.ts
import { describe, it, expect, vi } from 'vitest'
import { OfficeMockObject } from 'office-addin-mock'
import { yourFunction } from '../app/src/taskpane/{slug}'

describe('Problem N: {Title}', () => {
  it('{first observable behaviour}', async () => {
    const mockData = {
      context: {
        document: {
          body: {
            text: 'Initial body text',
            load: vi.fn(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      },
      run: async (callback) => callback(mockData.context),
    }
    const mock = new OfficeMockObject(mockData)
    globalThis.Word = mock

    const result = await yourFunction()

    expect(result).toBe(...)
    expect(mockData.context.document.body.load).toHaveBeenCalledWith('text')
    expect(mockData.context.sync).toHaveBeenCalledTimes(1)
  })

  it('{second observable behaviour}', async () => {
    // ...
  })
})
```

## The verify script

A single shell script at the repo root:

```bash
#!/usr/bin/env bash
# verify — usage: ./verify N
set -euo pipefail

N=$(printf "%02d" "$1")
TEST_FILE=$(ls tests/${N}-*.test.ts 2>/dev/null | head -1)

if [[ -z "$TEST_FILE" ]]; then
  echo "No test found for problem $N"
  exit 1
fi

cd app
echo "→ typecheck"
npm run typecheck

echo "→ vitest"
cd ..
npx vitest run "$TEST_FILE"

echo ""
echo "✓ Automated verification passed for problem $N"
echo ""
echo "─── MANUAL CHECKLIST ───"
echo "Sideload app/manifest.xml into Word Web or Word Desktop, then confirm:"
echo ""

# Print the manual checklist section from the problem file
PROBLEM_FILE=$(ls problems/${N}-*.md 2>/dev/null | head -1)
awk '/^## Manual verification/,/^## Hints/' "$PROBLEM_FILE" | grep '^\- \[ \]'

echo ""
echo "Confirm each item yourself. When all confirmed, problem $N is done."
```

## What the generator must guarantee

For each problem N:

1. `tests/N-{slug}.test.ts` exists and follows the file shape above.
2. The vitest test FAILS against the `main` branch (where `app/` has only previous problems' code).
3. The vitest test PASSES against `solutions/N`.
4. The manual checklist in `problems/N-*.md` is unambiguous and confirmable in <2 minutes by a competent learner.
5. Solution branches stay isolated (`main` never contains solutions).
