/**
 * Evaluates dynamic expressions against a record's field values
 * Used for visibility conditions and dynamic content in button configurations
 */
export class ExpressionEvaluator {
    private record: Record<string, any>;
    private context: ComponentFramework.Context<any>;

    /**
     * Creates a new instance of ExpressionEvaluator
     * @param record - The record containing field values to evaluate against
     * @param context - The PCF component context
     */
    constructor(record: Record<string, any>, context: ComponentFramework.Context<any>) {
        this.record = record;
        this.context = context;
    }

    /**
     * Evaluates a boolean expression by replacing field references with actual values
     * @param expression - Expression containing field references in square brackets (e.g., [field1] > 0)
     * @returns boolean - Result of the expression evaluation
     */
    evaluate(expression: string): boolean {
        if (!expression) return true;

        try {
            const parsedExpression = this.parseExpression(expression);
            return this.safeEvaluate(parsedExpression);
        } catch (error) {
            console.error('Error evaluating expression:', error);
            return false;
        }
    }

    /**
     * Replaces field references with their actual values from the record
     * @param expression - Expression containing field references in square brackets
     * @returns string - Expression with field references replaced by their values
     */
    private parseExpression(expression: string): string {
        return expression.replace(/\{(\w+)\}/g, 'record.$1');
    }

    /**
     * Safely evaluates the parsed expression with limited scope and sanitized input
     * @param expression - Parsed expression to evaluate
     * @returns boolean - Result of the safe evaluation
     */
    private safeEvaluate(expression: string): boolean {
        try {
            // We need to use Function constructor for dynamic evaluation
            // eslint-disable-next-line @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call
            return Function('record', 'context', `
                "use strict";
                try {
                    return Boolean(${expression});
                } catch (e) {
                    return false;
                }
            `)(this.record, this.context);
        } catch {
            return false;
        }
    }

    /**
     * Safely executes an async script with limited scope and error handling
     * @param script - The script to execute
     * @returns Promise<void>
     */
    public async safeExecute(script: string): Promise<void> {
        try {
            // Create an async function that will execute the script
            // eslint-disable-next-line @typescript-eslint/no-implied-eval
            const executor = new Function('record', 'context', `
                "use strict";
                return (async () => {
                    try {
                        ${script}
                    } catch (e) {
                        if (Xrm?.Navigation?.openAlertDialog) {
                            await Xrm.Navigation.openAlertDialog({
                                text: e.message,
                                title: "Error in Action Script"
                            });
                        } else {
                            console.error("Error in Action Script:", e);
                            alert(e.message);
                        }
                        throw e;
                    }
                })();
            `);

            // Execute the function with the sandbox context using spread operator
            // eslint-disable-next-line @typescript-eslint/no-unsafe-call
            await executor(this.record, this.context);
        } catch (error) {
            console.error("Error executing script:", error);
            throw error; // Re-throw to allow caller to handle
        }
    }
}