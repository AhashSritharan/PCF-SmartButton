/* eslint-disable @typescript-eslint/no-unsafe-call */
import * as React from "react";
import { DefaultButton, TooltipHost, Link, FontIcon, MessageBar, MessageBarType } from "@fluentui/react";
import { ExpressionEvaluator } from "./ExpressionEvaluator";

// Add types to window object for global record cache
type RecordCache = Record<string, {
  record?: Record<string, any>;
  promise?: Promise<Record<string, any>>;
}>;

(window as any).recordCache = (window as any).recordCache || {};

/**
 * Configuration interface for Smart Button control
 * Defines the structure of button settings stored in Dataverse
 */
export interface ButtonConfig {
  theia_buttonlabel: string;
  theia_url: string;
  theia_buttonposition: number;
  theia_buttontooltip?: string;
  theia_tablename: string;
  theia_buttonicon?: string;
  theia_visibilityexpression?: string;
  theia_showaslink: boolean;
  theia_actionscript?: string;
  isVisible?: boolean;  // Added to support visibility state
}

/**
 * Props for the SmartButton component
 */
interface SmartButtonProps {
  configs: ButtonConfig[];
  context: ComponentFramework.Context<any>;
  recordId: string;
  entityName: string;
}

/**
 * Props for the ButtonRenderer component
 */
interface ButtonRendererProps {
  config: ButtonConfig;  // Use the full ButtonConfig type since we're passing the entire config object
  onClick: () => void;
}

// Define the component first, then memoize it
const ButtonRendererBase: React.FC<ButtonRendererProps> = ({ config, onClick }) => {
  return (
    config.theia_showaslink ? (
      <Link href={config.theia_url} target="_blank" rel="noopener noreferrer">
        {config.theia_buttonlabel}
      </Link>
    ) : (
      <DefaultButton
        onClick={onClick}
        styles={{
          root: {
            minWidth: 'auto',
            padding: '0 12px',
            margin: 0
          }
        }}
      >
        {config.theia_buttonicon && (
          <FontIcon
            iconName={config.theia_buttonicon}
            style={{ marginRight: 8, fontSize: 16 }}
          />
        )}
        <span>{config.theia_buttonlabel}</span>
      </DefaultButton>
    )
  );
};

const ButtonRenderer = React.memo(ButtonRendererBase);

/**
 * Error boundary component to catch and handle rendering errors gracefully
 * Prevents the entire control from crashing when a button fails to render
 */
class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { hasError: boolean }
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(): { hasError: boolean } {
    return { hasError: true };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    console.error('Error caught by boundary:', error, errorInfo);
  }

  render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          An error occurred while rendering the buttons. Please refresh the page.
        </MessageBar>
      );
    }

    return this.props.children;
  }
}

/**
 * Main SmartButton component that manages button configurations and rendering
 * Features:
 * - Dynamic button visibility based on record field values
 * - Caching of record data for performance
 * - Support for related record field references
 * - Custom action script execution
 */

export const SmartButton: React.FC<SmartButtonProps> = ({ configs, context, recordId, entityName }) => {
  /**
   * Component state
   */
  const [record, setRecord] = React.useState<Record<string, any> | null>(null);
  const [isLoading, setIsLoading] = React.useState(true);
  const [error, setError] = React.useState<string | null>(null);
  const [expressionEvaluator, setExpressionEvaluator] = React.useState<ExpressionEvaluator | null>(null);
  const [resolvedConfigs, setResolvedConfigs] = React.useState<ButtonConfig[]>([]);
  const [saveCounter, setSaveCounter] = React.useState(0); // Add counter to trigger re-renders

  // Add a ref to track all fetched record keys
  const fetchedRecordKeys = React.useRef<Set<string>>(new Set());

  // Add save event handler
  React.useEffect(() => {
    const handlePostSave = () => {
      const cache = (window as any).recordCache as Record<string, CacheEntry>;
      fetchedRecordKeys.current.forEach(key => {
        delete cache[key];
      });
      setSaveCounter(prev => prev + 1);
    };

    try {
      const formContext = (window as any).Xrm?.Page?.data?.entity;
      if (formContext?.addOnPostSave) {
        const handler = () => handlePostSave();
        formContext.addOnPostSave(handler);
        return () => {
          if (formContext.removeOnPostSave) {
            formContext.removeOnPostSave(handler);
          }
        };
      }
    } catch (error) {
      console.warn('Failed to add post save handler:', error);
    }
  }, []); // Empty dependency array as we only want to set up the handler once

  // Add cleanup for cache on unmount with complete record tracking
  React.useEffect(() => {
    return () => {
      const cache = (window as any).recordCache as Record<string, CacheEntry>;
      fetchedRecordKeys.current.forEach(key => {
        delete cache[key];
      });
    };
  }, []);

  /**
   * Type definition for cached records
   */
  type CachedRecord = Record<string, any>;

  interface CacheEntry {
    record?: Record<string, any>;
    isLoading?: boolean;
  }

  /**
   * Enhanced record retrieval with global caching and request deduplication
   */
  const retrieveRecordWithCache = async (entityName: string, recordId: string, retryCount = 0): Promise<CachedRecord> => {
    const MAX_RETRIES = 5;
    const RETRY_DELAY = 100; // milliseconds
    const cacheKey = `${entityName}:${recordId}`;
    const cache = (window as any).recordCache as Record<string, CacheEntry>;

    // Track this record key immediately when the request is made
    if (!fetchedRecordKeys.current.has(cacheKey)) {
      fetchedRecordKeys.current.add(cacheKey);
    }

    if (cache[cacheKey]?.record) {
      return cache[cacheKey].record!;
    }

    if (cache[cacheKey]?.isLoading) {
      if (retryCount >= MAX_RETRIES) {
        delete cache[cacheKey];
      } else {
        try {
          await new Promise(resolve => setTimeout(resolve, RETRY_DELAY * (retryCount + 1)));
          return await retrieveRecordWithCache(entityName, recordId, retryCount + 1);
        } catch (error) {
          delete cache[cacheKey];
          throw error;
        }
      }
    }

    cache[cacheKey] = { isLoading: true };

    try {
      const record = await context.webAPI.retrieveRecord(entityName, recordId);
      cache[cacheKey] = { record };
      return record;
    } catch (error) {
      delete cache[cacheKey];
      throw error;
    }
  };

  /**
   * Fetches the main record data and initializes the expression evaluator
   */
  const fetchRecord = React.useCallback(async () => {
    try {
      const validRecordId = recordId && recordId !== "00000000-0000-0000-0000-000000000000";
      if (!validRecordId) {
        setRecord({});
        setExpressionEvaluator(new ExpressionEvaluator({}, context));
        return;
      }

      const result = await retrieveRecordWithCache(entityName, recordId);
      setRecord(result);
      setExpressionEvaluator(new ExpressionEvaluator(result, context));
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
      setError(errorMessage);
      setRecord({});
      setExpressionEvaluator(new ExpressionEvaluator({}, context));
    } finally {
      setIsLoading(false);
    }
  }, [context, entityName, recordId, saveCounter]); // Add saveCounter to dependencies

  // Fetch record effect
  React.useEffect(() => {
    void fetchRecord();
  }, [fetchRecord]);

  /**
   * Helper function to fetch a single level of related record
   */
  const fetchSingleLevelRecord = async (currentRecord: any, field: string): Promise<{ record: any; entity: string } | null> => {
    const idField = `_${field}_value`;
    const logicalNameField = `_${field}_value@Microsoft.Dynamics.CRM.lookuplogicalname`;

    const id = currentRecord[idField];
    const entity = currentRecord[logicalNameField];

    if (id && entity && typeof id === 'string' && typeof entity === 'string') {
      const cacheKey = `${entity}:${id}`;
      // Track this related record key
      fetchedRecordKeys.current.add(cacheKey);

      const record = await retrieveRecordWithCache(entity, id);
      return { record, entity };
    }
    return null;
  };

  /**
   * Fetches a related record when its fields are referenced in button configurations
   * Supports nested lookups by recursively fetching related records
   */
  const fetchRelatedRecord = async (lookupPath: string[]): Promise<Record<string, any> | null> => {
    if (!record || lookupPath.length === 0) return null;

    try {
      let currentRecord = record;
      let result: Record<string, any> | null = null;

      // Handle first level
      const firstLevel = await fetchSingleLevelRecord(currentRecord, lookupPath[0]);
      if (!firstLevel) return null;

      result = firstLevel.record;
      currentRecord = firstLevel.record;

      // Handle additional levels if they exist
      for (let i = 1; i < lookupPath.length; i++) {
        const nextLevel = await fetchSingleLevelRecord(currentRecord, lookupPath[i]);
        if (!nextLevel) break;

        currentRecord = nextLevel.record;
        result = nextLevel.record;
      }

      return result;
    } catch (error) {
      console.error('Error retrieving related record for', lookupPath.join('.'), error);
      return null;
    }
  };

  /**
   * Replaces dynamic field references in text with actual values
   * Supports both direct and nested lookup field references
   */
  const asyncReplaceDynamicValues = async (text: string | undefined): Promise<string> => {
    if (!text || !record) return '';
    const regex = /\{([^}]+)\}/g;
    let newText = text;
    let match: RegExpExecArray | null;

    while ((match = regex.exec(text)) !== null) {
      const [fullMatch, token] = match;
      let value;

      if (token.includes('.')) {
        const parts = token.split('.');
        const lookupPath = parts.slice(0, -1);
        const finalField = parts[parts.length - 1];

        // Special handling for ID field
        if (finalField.toLowerCase() === 'id') {
          const lookupIdField = `_${lookupPath[0]}_value`;
          value = record[lookupIdField] ?? fullMatch;
        } else {
          // For other fields, fetch the related records through the lookup chain
          const relatedRecord = await fetchRelatedRecord(lookupPath);
          value = relatedRecord?.[finalField] ?? fullMatch;
        }
      } else {
        value = record[token] ?? fullMatch;
      }

      newText = newText.replace(fullMatch, String(value));
    }

    return newText;
  };

  /**
   * Processes button configurations:
   * - Filters based on visibility expressions
   * - Sorts by position
   * - Resolves dynamic values
   */
  React.useEffect(() => {
    if (!record) return;

    const resolveConfigs = async () => {
      try {
        const sortedConfigs = [...configs].sort((a, b) => a.theia_buttonposition - b.theia_buttonposition);
        const newConfigs: ButtonConfig[] = [];

        // Process configs sequentially to prevent parallel record retrievals
        for (const config of sortedConfigs) {
          const resolvedConfig = {
            ...config,
            theia_buttonlabel: await asyncReplaceDynamicValues(config.theia_buttonlabel),
            theia_buttontooltip: await asyncReplaceDynamicValues(config.theia_buttontooltip),
            theia_url: await asyncReplaceDynamicValues(config.theia_url),
            theia_actionscript: await asyncReplaceDynamicValues(config.theia_actionscript),
            theia_visibilityexpression: await asyncReplaceDynamicValues(config.theia_visibilityexpression)
          };

          // Only include buttons that pass visibility check
          const isVisible = !resolvedConfig.theia_visibilityexpression ||
            expressionEvaluator?.evaluate(resolvedConfig.theia_visibilityexpression);

          if (isVisible) {
            newConfigs.push({ ...resolvedConfig, isVisible });
          }
        }

        setResolvedConfigs(newConfigs);
      } catch (error) {
        console.error('Error resolving configs:', error);
        setError('Failed to process button configurations');
      }
    };

    void resolveConfigs();
  }, [record, configs, expressionEvaluator]);

  /**
   * Handles button clicks:
   * - Executes custom action scripts if provided
   * - Opens URLs in new tab otherwise
   */
  const handleClick = async (url: string | undefined, actionScript: string | undefined) => {
    if (actionScript && expressionEvaluator) {
      try {
        const parsedScript = await asyncReplaceDynamicValues(actionScript);
        await expressionEvaluator.safeExecute(parsedScript);
      } catch (error) {
        console.error('Error executing action script:', error);
      }
    } else if (url) {
      window.open(url, '_blank', 'noopener,noreferrer');
    }
  };

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  if (isLoading) {
    return <MessageBar messageBarType={MessageBarType.info}>Loading buttons...</MessageBar>;
  }

  return (
    <ErrorBoundary>
      <div style={{
        display: 'flex',
        flexWrap: 'wrap',
        gap: '8px',
        padding: '4px'
      }}>
        {resolvedConfigs.map((config, index) => (
          <TooltipHost
            content={config.theia_buttontooltip}
            key={`${config.theia_buttonlabel}_${index}`}
          >
            <ButtonRenderer
              config={config}
              onClick={() => {
                void handleClick(config.theia_url, config.theia_actionscript);
              }}
            />
          </TooltipHost>
        ))}
      </div>
    </ErrorBoundary>
  );
};