/**
 * L3: reaching the document model from an operation, and saying the same thing
 * however the read fails.
 *
 * The model is loaded in two steps, and only the first of them is eager.
 * `WorkbookModel.load` opens the package; a worksheet part is parsed the first
 * time somebody asks for it, so a malformed row or cell reference surfaces from
 * an ordinary property access, long after the load appeared to succeed.
 * Wrapping only the load left half the read untranslated: one broken workbook
 * raised a stable `XLSX_READ_FAILED` naming the file, or a bare parser `Error`
 * naming nothing, depending on which part of it was broken.
 *
 * So both steps go through one translation here, and the words a failure is
 * reported in travel with the model rather than beside it. An operation loads a
 * `WorkbookRead` and asks it for worksheets; there is no way to hold the model
 * without also holding what to call it, so the eager and lazy failures cannot
 * describe the same read differently.
 */
import { ConsultChimpsError } from "@consultchimps/core";

import { XLSX_ERRORS } from "../errors.js";
import { WorkbookModel } from "../model/index.js";
import type { WorksheetModel } from "../model/types.js";

/**
 * How one read of one workbook is named when it fails: the workbook as the
 * caller knows it, and the details every error from this read carries.
 */
export interface WorkbookReadContext {
  /** The workbook's name in messages: a file path, or an input name. */
  readonly source: string;
  /** The details each surface identifies its input by. */
  readonly details: Record<string, unknown>;
}

/**
 * A loaded workbook, with the context its failures are reported in.
 *
 * Operations take this rather than a bare `WorkbookModel` wherever they force a
 * worksheet to parse. Everything that only reads an already-parsed worksheet
 * keeps taking the model itself.
 */
export class WorkbookRead {
  readonly workbook: WorkbookModel;
  readonly context: WorkbookReadContext;

  private constructor(workbook: WorkbookModel, context: WorkbookReadContext) {
    this.workbook = workbook;
    this.context = context;
  }

  /**
   * Load the document model, reporting an unreadable package as the stable read
   * error every workbook reader raises.
   */
  static async load(
    bytes: Uint8Array,
    context: WorkbookReadContext,
  ): Promise<WorkbookRead> {
    try {
      return new WorkbookRead(await WorkbookModel.load(bytes), context);
    } catch (error) {
      throw readFailure(context, undefined, error);
    }
  }

  /**
   * The parsed worksheet, or undefined when the workbook holds no sheet by that
   * name.
   *
   * Absent and unreadable are different answers and stay different: a caller
   * may reasonably describe a missing sheet as empty, and none of them may
   * treat a worksheet nothing could parse as one that held nothing.
   */
  worksheet(sheet: string): WorksheetModel | undefined {
    try {
      return this.workbook.worksheet(sheet);
    } catch (error) {
      throw readFailure(this.context, sheet, error);
    }
  }
}

/**
 * The one read failure. A worksheet name joins the message and the details when
 * the failure belongs to a worksheet, and the cause is always carried, so a
 * caller reporting the error can say which workbook, which sheet, and what the
 * parser objected to.
 */
function readFailure(
  context: WorkbookReadContext,
  worksheet: string | undefined,
  cause: unknown,
): ConsultChimpsError {
  return new ConsultChimpsError(
    XLSX_ERRORS.XLSX_READ_FAILED,
    worksheet === undefined
      ? `Could not read workbook: ${context.source}`
      : `Could not read worksheet "${worksheet}" in workbook: ${context.source}`,
    {
      cause,
      details:
        worksheet === undefined
          ? context.details
          : { ...context.details, worksheet },
    },
  );
}
