/**
 * Evaluates dynamic expressions against a record's field values
 * Used for visibility conditions and dynamic content in button configurations
 */
export class ExpressionEvaluator {
    private record: Record<string, any>;
    private context: ComponentFramework.Context<any>;
    private relatedRecordsCache = new Map<string, Record<string, any>>();

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
     * @param expression - Expression containing field references in curly braces (e.g., {field1} > 0)
     * @returns boolean - Result of the expression evaluation
     */
    evaluate(expression: string): boolean {
        if (!expression) return true;

        try {
            const fieldReferences = this.extractFieldReferences(expression);
            if (fieldReferences.some(ref => ref.includes('.'))) {
                // If expression contains related field references, we can't evaluate it synchronously
                console.warn('Expression contains related fields and cannot be evaluated synchronously:', expression);
                return false;
            }

            // Create a data object with field values for simple references
            const data = this.createDataObject(fieldReferences);
            return this.safeEvaluateWithData(expression, data);
        } catch (error) {
            console.error('Error evaluating expression:', error);
            return false;
        }
    }

    /**
     * Evaluates a boolean expression asynchronously to support related record lookups
     * @param expression - Expression containing field references in curly braces (e.g., {field1} > 0)
     * @returns Promise<boolean> - Result of the expression evaluation
     */
    async evaluateAsync(expression: string): Promise<boolean> {
        if (!expression) return true;

        try {
            // Extract all field references from the expression
            const fieldReferences = this.extractFieldReferences(expression);

            // Create a data object with all field values, including related records
            const data = await this.createDataObjectAsync(fieldReferences);

            // Evaluate the expression with the complete data object
            return this.safeEvaluateWithData(expression, data);
        } catch (error) {
            console.error('Error evaluating async expression:', error);
            return false;
        }
    }

    /**
     * Extract all field references from an expression
     * @param expression - Expression containing field references in curly braces
     * @returns string[] - Array of field references
     */
    private extractFieldReferences(expression: string): string[] {
        const regex = /\{([^}]+)\}/g;
        const references: string[] = [];
        let match;

        regex.lastIndex = 0; // Reset regex index

        while ((match = regex.exec(expression)) !== null) {
            references.push(match[1]);
        }

        return references;
    }    /**
     * Creates a data object with field values for evaluation
     * @param fieldReferences - Array of field references to extract
     * @returns Record<string, any> - Object with field values mapped to paths
     */
    private createDataObject(fieldReferences: string[]): Record<string, any> {
        const data: Record<string, any> = {};

        for (const reference of fieldReferences) {
            if (!reference.includes('.')) {
                // Only handle direct references in the synchronous version
                data[reference] = this.convertValueToProperType(this.record[reference]);
            }
        }

        return data;
    }

    /**
     * Checks if a string appears to be a date format
     * @param str - String to check
     * @returns boolean - True if string looks like a date
     */
    private isDateString(str: string): boolean {
        // Check for ISO format dates or common Dataverse date formats
        const isoDatePattern = /^\d{4}-\d{2}-\d{2}(T|\s)\d{2}:\d{2}:\d{2}(Z|[+-]\d{2}:\d{2})?$/;
        return isoDatePattern.test(str);
    }


    /**
     * Creates a data object with all field values, including those from related records
     * @param fieldReferences - Array of field references to extract
     * @returns Promise<Record<string, any>> - Object with field values mapped to paths
     */
    private async createDataObjectAsync(fieldReferences: string[]): Promise<Record<string, any>> {
        const data: Record<string, any> = {};

        // Process all field references
        for (const reference of fieldReferences) {
            if (reference.includes('.')) {
                // Handle related record field references
                const value = await this.getRelatedFieldValue(reference);
                data[reference] = this.convertValueToProperType(value);
            } else {
                // Handle direct fields on main record
                data[reference] = this.convertValueToProperType(this.record[reference]);
            }
        }

        return data;
    }

    /**
     * Converts a field value to its proper JavaScript type, particularly for dates
     * @param value - The field value to convert
     * @returns any - The converted value with proper type
     */
    private convertValueToProperType(value: any): any {
        if (value === null || value === undefined) {
            return null;
        }

        // Convert date strings to Date objects
        if (typeof value === 'string' && this.isDateString(value)) {
            return new Date(value);
        }

        // Already a Date object, return as is
        if (value instanceof Date) {
            return value;
        }

        // For all other types, return as is
        return value;
    }

    /**
     * Gets a field value from a related record using dot notation path
     * @param path - Dot notation path to the field (e.g., "contact.fullname")
     * @returns Promise<any> - The resolved field value
     */
    private async getRelatedFieldValue(path: string): Promise<any> {
        if (!this.record) return null;

        const parts = path.split('.');
        const lookupPath = parts.slice(0, -1);
        const finalField = parts[parts.length - 1];

        try {
            // Check if we have this path cached
            const cacheKey = lookupPath.join('.');
            let relatedRecord = this.relatedRecordsCache.get(cacheKey);

            if (!relatedRecord) {
                // If not cached, retrieve the related record
                relatedRecord = (await this.fetchRelatedRecord(lookupPath)) ?? undefined;
                if (relatedRecord) {
                    this.relatedRecordsCache.set(cacheKey, relatedRecord);
                }
            }

            // Return the field value or null if not found
            return relatedRecord && relatedRecord[finalField] !== undefined
                ? relatedRecord[finalField]
                : null;
        } catch (error) {
            console.error('Error resolving related field:', path, error);
            return null;
        }
    }    /**
     * Fetches a related record by traversing lookup references
     * @param lookupPath - Array of lookup attribute names to traverse
     * @returns Promise<Record<string, any> | null> - The related record or null if not found
     */
    private async fetchRelatedRecord(lookupPath: string[]): Promise<Record<string, any> | null> {
        if (!this.record || lookupPath.length === 0) return null;

        try {
            const currentRecord = this.record;
            let currentEntity = '';
            let currentId = '';

            // Handle first level lookup
            const firstLookup = lookupPath[0];
            const idField = `_${firstLookup}_value`;
            const logicalNameField = `_${firstLookup}_value@Microsoft.Dynamics.CRM.lookuplogicalname`;

            currentId = currentRecord[idField];
            currentEntity = currentRecord[logicalNameField];

            if (!currentId || !currentEntity) return null;

            // Fetch the first level record
            let result = await this.retrieveRecord(currentEntity, currentId);
            if (!result) return null;

            // Handle additional levels if they exist
            for (let i = 1; i < lookupPath.length; i++) {
                const nextLookup = lookupPath[i];
                const nextIdField = `_${nextLookup}_value`;
                const nextLogicalNameField = `_${nextLookup}_value@Microsoft.Dynamics.CRM.lookuplogicalname`;

                currentId = result[nextIdField];
                currentEntity = result[nextLogicalNameField];

                if (!currentId || !currentEntity) return null;

                // Fetch the next level record
                const nextRecord = await this.retrieveRecord(currentEntity, currentId);
                if (!nextRecord) return null;

                result = nextRecord;
            }

            return result;
        } catch (error) {
            console.error('Error fetching related record:', lookupPath.join('.'), error);
            return null;
        }
    }

    /**
     * Retrieves a record from Dataverse with caching
     * @param entityName - Logical name of the entity
     * @param id - GUID of the record
     * @returns Promise<Record<string, any> | null> - The retrieved record or null
     */
    private async retrieveRecord(entityName: string, id: string): Promise<Record<string, any> | null> {
        try {
            return await this.context.webAPI.retrieveRecord(entityName, id);
        } catch (error) {
            console.error(`Error retrieving ${entityName} record with ID ${id}:`, error);
            return null;
        }
    }    /**
     * Safely evaluates the expression using direct field values (not string replacements)
     * @param expression - Original expression with field references in curly braces
     * @param data - Object containing field values mapped to their reference paths
     * @returns boolean - Result of the safe evaluation
     */
    private safeEvaluateWithData(expression: string, data: Record<string, any>): boolean {
        try {
            // Process the expression to handle both field references and method calls on those references
            const processedExpression = this.processExpressionWithMethodCalls(expression, data);

            // Create a function with data as a parameter
            // eslint-disable-next-line @typescript-eslint/no-implied-eval, @typescript-eslint/no-unsafe-call
            return Function('data', 'record', 'context', `
                "use strict";
                try {
                    return Boolean(${processedExpression});
                } catch (e) {
                    console.error('Expression evaluation error:', e);
                    return false;
                }
            `)(data, this.record, this.context);
        } catch (error) {
            console.error('Safe evaluate error:', error);
            return false;
        }
    }    /**
     * Processes an expression to handle both field references and method calls on field references
     * @param expression - Original expression with field references in curly braces
     * @param data - The data object containing field values
     * @returns string - Processed expression with field references replaced
     */
    private processExpressionWithMethodCalls(expression: string, data: Record<string, any>): string {
        // First, handle simple field references that don't have method calls
        // This handles cases like {fieldName} in expressions
        let processedExpression = expression.replace(/\{([^}]+)\}/g, (fullMatch, path) => {
            return `data['${path}']`;
        });

        // Now fix any method chains that might have been broken in the replacement
        // This regex looks for data['path'] followed by one or more method calls
        const methodChainRegex = /data\['([^']+)'\]((?:\.[a-zA-Z0-9_]+(?:\([^()]*\))?)+)/g;

        // Replace with the correct syntax that preserves the entire method chain
        processedExpression = processedExpression.replace(methodChainRegex, (fullMatch, path, methodChain) => {
            // For method chains, make sure the value exists before attempting to call methods
            return `(data['${path}'] != null ? data['${path}']${methodChain} : null)`;
        });

        return processedExpression;
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