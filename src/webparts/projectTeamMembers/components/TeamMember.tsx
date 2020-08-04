import * as React from 'react';
import { Facepile, IFacepilePersona } from 'office-ui-fabric-react/lib/Facepile';
import { PersonaSize, personaSize } from 'office-ui-fabric-react/lib/Persona';
import { IEnsureUser } from '../../../common/services/PeoplePickerService';
import { getInitials } from 'office-ui-fabric-react/lib/Utilities';
import styles from './ProjectTeamMembers.module.scss';

export interface ITeamMemberProps {
  user: IEnsureUser;
}

const TeamMember: React.FunctionComponent<ITeamMemberProps> = (props) => {
  const { user } = props;
  const facepile = { personaName: user.Title, imageInitials: getInitials(user.Title, false), imageUrl: '' };

  return (
    <Facepile className={styles.teamMember} personaSize={PersonaSize.size48} personas={[facepile]} />
  )
}

export default TeamMember;