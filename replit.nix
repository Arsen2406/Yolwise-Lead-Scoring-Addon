{ pkgs }: {
  deps = [
    pkgs.python311Full
    pkgs.python311Packages.pip
    pkgs.python311Packages.flask
    pkgs.python311Packages.werkzeug
  ];
}
