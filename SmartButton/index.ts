import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { SmartButton, ButtonConfig } from "./SmartButton";
import * as React from "react";

/**
 * Context information provided by the PCF framework
 */
interface ContextInfo {
    entityId: string;
    entityTypeName: string;
}

/**
 * Props for the ButtonsWrapper component
 */
interface ButtonsWrapperProps {
    context: ComponentFramework.Context<IInputs>;
    recordId: string;
    entityName: string;
    buttonFilter: string;
}

/**
 * Custom hook to fetch and manage button configurations
 * @param context - PCF context
 * @param entityName - Name of the current entity
 * @param buttonFilter - Optional filter for button configurations
 * @returns Object containing button configs, loading state, and error state
 */
const useFetchButtonConfigs = (
    context: ComponentFramework.Context<IInputs>,
    entityName: string,
    buttonFilter: string
) => {
    const [buttonConfigs, setButtonConfigs] = React.useState<ButtonConfig[]>([]);
    const [isLoading, setIsLoading] = React.useState(true);
    const [error, setError] = React.useState<string | null>(null);

    React.useEffect(() => {
        const fetchConfigs = async () => {
            try {
                let filter = `theia_tablename eq '${entityName}' and statecode eq 0`;
                if (buttonFilter) {
                    filter += ` and ${buttonFilter}`;
                }

                const query = `?$select=theia_buttonlabel,theia_url,theia_buttonposition,theia_buttontooltip,theia_tablename,theia_buttonicon,theia_visibilityexpression,theia_showaslink,theia_actionscript&$filter=${filter}&$orderby=theia_buttonposition asc`;

                const result = await context.webAPI.retrieveMultipleRecords(
                    "theia_buttonconfiguration",
                    query
                );

                setButtonConfigs(result.entities as ButtonConfig[]);
            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : 'Failed to fetch button configurations';
                setError(errorMessage);
                setButtonConfigs([]);
            } finally {
                setIsLoading(false);
            }
        };

        void fetchConfigs();
    }, [context.webAPI, entityName, buttonFilter]);

    return { buttonConfigs, isLoading, error };
};

/**
 * Wrapper component that handles button configuration fetching and error states
 */
const ButtonsWrapper: React.FC<ButtonsWrapperProps> = ({
    context,
    recordId,
    entityName,
    buttonFilter
}) => {
    const { buttonConfigs, isLoading, error } = useFetchButtonConfigs(
        context,
        entityName,
        buttonFilter
    );

    if (error) {
        return React.createElement("div", { role: "alert" }, `Error: ${error}`);
    }

    if (isLoading) {
        return React.createElement("div", { role: "status" }, "Loading buttons...");
    }

    return React.createElement(SmartButton, {
        configs: buttonConfigs,
        context,
        recordId,
        entityName
    });
};

/**
 * Main PCF control class implementing the smart button functionality
 * Features:
 * - Dynamic button rendering based on configurations
 * - Support for visibility conditions
 * - Caching for performance optimization
 * - Error handling and loading states
 */
export class SmartButtonControl implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private recordId: string;
    private entityName: string;
    private buttonFilter: string;

    /**
     * Initializes the control with configuration data from the framework
     * @param context - PCF context containing parameters and configuration
     * @param notifyOutputChanged - Callback to notify framework of changes
     * @param state - Current state
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        const contextInfo = (context.mode as any).contextInfo as ContextInfo;

        this.recordId = contextInfo?.entityId ?? "";
        this.entityName = contextInfo?.entityTypeName ?? "";
        this.buttonFilter = context.parameters?.ButtonFilter?.raw ?? "";
    }

    /**
     * Updates the control when framework detects a change in bound properties
     * @param context - Updated PCF context
     * @returns React element to render
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.buttonFilter = context.parameters?.ButtonFilter?.raw ?? "";

        return React.createElement(ButtonsWrapper, {
            context,
            recordId: this.recordId,
            entityName: this.entityName,
            buttonFilter: this.buttonFilter
        });
    }

    /**
     * Cleanup handler called when the control is removed from the form
     */
    public destroy(): void {
        // Add cleanup code here
    }

    /**
     * Returns control outputs to the framework
     * Currently not used as this is a display-only control
     */
    public getOutputs(): IOutputs {
        return {};
    }
}