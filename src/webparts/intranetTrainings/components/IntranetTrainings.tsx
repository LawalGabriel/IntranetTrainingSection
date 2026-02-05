/* eslint-disable react/self-closing-comp */
/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { useState, useEffect, useRef, useCallback } from 'react';
import styles from './IntranetTrainings.module.scss';
import type { IIntranetTrainingsProps, ITrainingItems } from './IIntranetTrainingsProps';
import { SPFx, spfi } from '@pnp/sp';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Placeholder } from '@pnp/spfx-controls-react';

const IntranetTrainings: React.FC<IIntranetTrainingsProps> = (props) => {
  const [trainingItems, setTrainingItems] = useState<ITrainingItems[]>([]);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const spRef = useRef<any>(null);

  // Format date like "29 DEC 2025"
  const formatDate = (dateString: string): string => {
    if (!dateString) return "";
    
    try {
      const date = new Date(dateString);
      // Check if date is valid
      if (isNaN(date.getTime())) {
        return "";
      }
      
      const day = date.getDate();
      const month = date.toLocaleString('en-US', { month: 'short' }).toUpperCase();
      const year = date.getFullYear();
      return `${day} ${month} ${year}`;
    } catch (error) {
      console.error('Error formatting date:', error);
      return "";
    }
  };

  // Get display date - use ScheduledDate if available and configured, otherwise use Created
  const getDisplayDate = (item: ITrainingItems): string => {
    if (props.showScheduledDate && item.ScheduledDate) {
      return formatDate(item.ScheduledDate);
    }
    return formatDate(item.Created);
  };

  const loadTrainingItems = useCallback(async () => {
    try {
      setIsLoading(true);
      setErrorMessage('');

      const items: ITrainingItems[] = await spRef.current.web.lists
        .getByTitle(props.listTitle)
        .items
        .select(
          "Id",
          "Title",
          "Link",
          "AttachmentFiles",
          "Status",
          "Created",
          "ScheduledDate",
          "Category"
        )
        .expand("AttachmentFiles")
        .filter("Status eq 1")
        .orderBy("Created", false)();

      setTrainingItems(items);
      setIsLoading(false);

    } catch (error: any) {
      console.error('Error loading training items:', error);
      setIsLoading(false);
      setErrorMessage(`Failed to load training items. Please check if the list "${props.listTitle}" exists and you have permissions. Error: ${error.message}`);
    }
  }, [props.listTitle]);

  useEffect(() => {
    spRef.current = spfi().using(SPFx(props.context));
    void loadTrainingItems();
  }, [props.listTitle, props.context]);

  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <div className={styles.loadingSpinner}></div>
        <div>Loading training items...</div>
      </div>
    );
  }

  if (errorMessage) {
    return (
      <div className={styles.errorContainer}>
        <Placeholder
          iconName='Error'
          iconText='Error'
          description={errorMessage}
        >
          <button
            className={styles.retryButton}
            onClick={() => loadTrainingItems()}
          >
            Retry
          </button>
        </Placeholder>
      </div>
    );
  }



  return (
  <div className={`${styles.intranetTrainings} ${props.useFullWidth ? styles.fullWidth : ''}`}>
    <div className={styles.container}>
      {/* Centered Webpart Title */}
      {props.webPartTitle && (
        <div 
          className={styles.webpartTitleContainer}
          style={{ 
            backgroundColor: props.titleBackgroundColor || 'transparent',
            padding: props.titleBackgroundColor !== 'transparent' ? '1.5rem' : '0',
            borderRadius: props.titleBackgroundColor !== 'transparent' ? '8px' : '0',
            marginBottom: '2rem',
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center'
          }}
        >
          <h1 
            className={styles.webpartTitle}
            style={{ 
              color: props.titleFontColor || '#000000',
              fontWeight: props.titleFontWeight || '600',
              fontSize: `${props.titleFontSize || 32}px`,
              textAlign: 'center',
              margin: '0',
              paddingBottom: '0.5rem',
              borderBottom: props.titleBackgroundColor === 'transparent' ? '2px solid #e1e1e1' : 'none',
              width: props.titleBackgroundColor === 'transparent' ? '100%' : 'auto'
            }}
          >
            {props.webPartTitle}
          </h1>
        </div>
      )}

      {/* Training Items Grid */}
      <div 
        className={styles.trainingGrid}
        style={{
          gridTemplateColumns: `repeat(${props.itemsPerRow || 3}, 1fr)`
        }}
      >
        {trainingItems.length === 0 ? (
          <div className={styles.noItems}>
            <div>No training items found.</div>
          </div>
        ) : (
          trainingItems.map((item: ITrainingItems) => (
            <a
              key={item.Id}
              href={
                typeof item.Link === 'object' && item.Link !== null && 'Url' in item.Link
                  ? item.Link.Url
                  : (typeof item.Link === 'string' ? item.Link : '#')
              }
              target="_blank"
              rel="noopener noreferrer"
              className={styles.trainingCardLink}
            >
              <div 
                className={styles.trainingCard}
                style={{
                  backgroundColor: props.cardBackgroundColor || '#ffffff',
                  borderColor: props.cardBorderColor || '#e1e1e1',
                  minHeight: `${props.cardHeight || 100}px`
                }}
              >
                {/* Left Section - Scheduled Date (Conditional) */}
                {props.showScheduledDate && (
                  <div className={styles.cardLeft}>
                    <div 
                      className={styles.cardDate}
                      style={{ color: props.dateColor || '#333333' }}
                    >
                      {getDisplayDate(item)}
                    </div>
                  </div>
                )}

                {/* Middle Section - Title Only */}
                <div className={styles.cardMiddle}>
                  <h3 
                    className={styles.cardTitle}
                    style={{ color: props.titleColor || '#000000' }}
                  >
                    {item.Title}
                  </h3>
                </div>

                {/* Category Section - Small button on right side */}
                {props.showCategory && item.Category && (
                  <div className={styles.categoryContainer}>
                    <div 
                      className={styles.categoryButton}
                      style={{
                        backgroundColor: props.categoryBgColor || '#f0f0f0',
                        color: props.categoryColor || '#333333'
                      }}
                    >
                      <span className={styles.categoryText}>
                        {item.Category}
                      </span>
                    </div>
                  </div>
                )}
              </div>
            </a>
          ))
        )}
      </div>
    </div>
  </div>
);
};

export default IntranetTrainings;