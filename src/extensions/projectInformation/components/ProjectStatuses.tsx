import * as React from "react";
import styles from './ProjectInformation.module.scss';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import { Project } from "../../../common/model/Project";

const ProjectStatuses: React.FunctionComponent<{ project: Project }> = (props) => {
  const { project } = props;
  const statuses = [
    'Finance',
    'Scope',
    'Quality',
    'Team',
    'Relation',
    'Time'
  ];


  return (
    <div className={styles.projectStatuses}>
      {statuses.map((s) => (
        <Toggle className={styles[project[s + 'Status'].toLowerCase()]} checked={true} label={s.toUpperCase()}  />
      ))}
      
      {
        /* 
          Alternatively:
          <Toggle className={styles[s]} checked={true} styles={{ pill: { background: '#aaaaaa'} }} /> 
        */
      }
    </div>
  );
};

export default ProjectStatuses;