import React from "react";
import PropTypes from "prop-types";

const Button = ({ id, label, onClick }) => {
  return (
    <div role="button" id={id} className="ms-welcome__action ms-Button ms-Button--hero ms-font-xl" onClick={onClick}>
      <span className="ms-Button-label">{label}</span>
    </div>
  );
};

Button.propTypes = {
  id: PropTypes.string.isRequired,
  label: PropTypes.string.isRequired,
  onClick: PropTypes.func.isRequired,
};

export default Button;
