/**
 * Update many records as one step, reporting each refusal against its own
 * request.
 *
 * A spreadsheet gesture is not one write. A paste of a block, or a drag of the
 * fill handle, is tens or hundreds of cell writes that a person made as a
 * single movement, and they have to reach the file that way: one transaction,
 * so a crash or a failure part way through cannot leave a torn gesture behind,
 * and one round trip, so a caller in a worker is not posting a command per cell.
 *
 * The other half is what happens to a value the schema will not hold. Two
 * models were weighed. All-or-nothing discards every accepted write when one
 * cell is wrong, which for a two hundred cell paste means one bad value costs
 * the visitor the other hundred and ninety nine with nothing to show for it.
 * Partial application keeps them and reports the refusal against the cell it
 * belongs to, which is what a spreadsheet does and what the caller's own
 * per-cell explanation model already expects.
 *
 * So: partial application, atomically committed. Every request runs inside one
 * transaction; a request the library refuses is recorded and the batch carries
 * on; the accepted writes commit together. A failure that is not a refusal is a
 * fault nobody planned for, so it rolls the whole transaction back and
 * propagates, rather than committing the half of the gesture that happened to
 * run before it. An error saying the workspace itself is damaged counts as a
 * fault, not a refusal, even though the library raised it.
 */
import {
  isConsultChimpsError,
  type ConsultChimpsError,
} from "@consultchimps/core";
import type { CellValue } from "@consultchimps/tabular";

import type { Database } from "./database.js";

/** One record's worth of a batch: which record, and the columns to write. */
export interface RecordUpdateRequest {
  readonly table: string;
  /** The Record ID, which addresses the record and is never itself written. */
  readonly recordId: string;
  readonly values: Readonly<Record<string, CellValue>>;
}

/** A request the database accepted, and what it now holds for those columns. */
export interface AcceptedRecordUpdate {
  readonly accepted: true;
  /** The Record ID as it is stored, whatever casing the request used. */
  readonly recordId: string;
  /** The written columns, read back through their declared types. */
  readonly values: Record<string, CellValue>;
}

/**
 * A request the database refused, with the sentence and the stable code the
 * refusal carried. They travel as data rather than as the error object so a
 * caller across a worker boundary can rebuild and show them.
 */
export interface RefusedRecordUpdate {
  readonly accepted: false;
  /** The Record ID as the request gave it: there may be no stored one. */
  readonly recordId: string;
  readonly code: string;
  readonly message: string;
}

/** What became of one request. One per request, in the order they were given. */
export type RecordUpdateOutcome = AcceptedRecordUpdate | RefusedRecordUpdate;

/**
 * Apply a batch of record updates in one transaction.
 *
 * Each request goes through `Database.updateRecord`, so the conversion, the
 * Record ID rule, and the constraint translation are the single-record path's
 * own and cannot drift from it. An empty batch writes nothing and reports
 * nothing, which is what "update these zero records" means.
 */
export function updateRecords(
  database: Database,
  requests: readonly RecordUpdateRequest[],
): RecordUpdateOutcome[] {
  if (requests.length === 0) {
    return [];
  }
  return database.sql.transaction(() =>
    requests.map((request): RecordUpdateOutcome => {
      try {
        const updated = database.updateRecord(
          request.table,
          request.recordId,
          request.values,
        );
        return {
          accepted: true,
          recordId: updated.recordId,
          values: updated.values,
        };
      } catch (error) {
        // A refusal belongs to its request. Anything else is a fault, and
        // rethrowing it here is what takes the whole transaction back off.
        if (!isRefusal(error)) {
          throw error;
        }
        return refusalOf(request, error);
      }
    }),
  );
}

/**
 * Whether an error says the request was wrong, as opposed to the file.
 *
 * Every code this package raises for a damaged workspace begins `DB_CORRUPT_`:
 * a stored schema that does not match its tables, a stored value the column's
 * type cannot read back, a Record ID counter that has gone wrong. Such an
 * error is a library error, but it is about the workspace and not about the
 * request that happened to find it, and committing the requests beside it
 * would write into a file already known to be broken. So it is a fault here,
 * and the transaction goes back.
 */
function isRefusal(error: unknown): error is ConsultChimpsError {
  return isConsultChimpsError(error) && !error.code.startsWith("DB_CORRUPT_");
}

function refusalOf(
  request: RecordUpdateRequest,
  error: ConsultChimpsError,
): RefusedRecordUpdate {
  return {
    accepted: false,
    recordId: request.recordId,
    code: error.code,
    message: error.message,
  };
}
