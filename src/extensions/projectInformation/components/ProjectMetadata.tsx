import * as React from "react";
import styles from './ProjectInformation.module.scss';
import { Project } from "../../../common/model/Project";

export default function ProjectMetadata(props: { project: Project }) {
  return (
    <div className={styles.projectMetadata}>
      <div className={styles.projectTitle}>{props.project.Title}</div>
      <div className={styles.projectDetailContainer}>
        <div className={styles.projectDetail}>
          <div className={styles.detailLabel}>Client</div>
          <div className={styles.detailValue}>{props.project.Client}</div>
        </div>
        <div className={styles.projectDetail}>
          <div className={styles.detailLabel}>Project Number</div>
          <div className={styles.detailValue}>{props.project.ProjectNumber}</div>
        </div>
        <div className={styles.projectDetail}>
          <div className={styles.detailLabel}>Unit</div>
          <div className={styles.detailValue}>{props.project.Unit}</div>
        </div>
      </div>
    </div>
  );
}
