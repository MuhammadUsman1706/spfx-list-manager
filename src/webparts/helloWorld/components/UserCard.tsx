import * as React from "react";
import { FunctionComponent } from "react";
import {
  DefaultButton,
  PrimaryButton,
  ActionButton,
} from "@fluentui/react/lib/Button";

import styles from "./UserCard.module.scss";

const UserCard: FunctionComponent<any> = ({ user, deleteUser }) => {
  const deleteUserHandler = (event: any, id: string, etag: string): any => {
    event.preventDefault();
    deleteUser(id, etag);
  };

  return (
    <div className={styles.card}>
      <div className={styles.upper}>
        <div className={styles.image}>
          <img
            src="https://imgv3.fotor.com/images/blog-cover-image/10-profile-picture-ideas-to-make-you-stand-out.jpg"
            alt="p_pic"
          />
        </div>
        <h1>{user.Title}</h1>
        <p>{user.Role}</p>
      </div>
      <div className={styles.lower}>
        <div>
          <h3>241</h3>
          <p>Photos</p>
        </div>
        <div>
          <h3>84K</h3>
          <p>Followers</p>
        </div>
      </div>
      <ActionButton
        onClick={(event) =>
          deleteUserHandler(event, user.ID, user["@odata.etag"])
        }
        className={styles["delete-button"]}
      >
        Delete
      </ActionButton>
    </div>
  );
};

export default UserCard;
