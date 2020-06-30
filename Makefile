SHELL=/bin/bash
.DEFAULT_GOAL:=help


.PHONY: help info
##@ Helpers

help:  ## Display this help
	@awk 'BEGIN {FS = ":.*##"; printf "\nUsage:\n  \
	make [VARS] \033[36m<target>\033[0m\n"} /^[a-zA-Z_-]+:.*?##/ \
	{ printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2 } /^##@/ \
	{ printf "\n\033[1m%s\033[0m\n", substr($$0, 5) } ' $(MAKEFILE_LIST)

################################################################################

B231 := 10.102.15.231
BUS_UI := ui.dev.axpl.prodv.net
S4 := 10.102.15.47
ASMA_PROD := asma.prodv.net
ASMA_STAGE := 10.54.254.12
LOYAL := ui.stage.loyal.prodv.net
PATN := ui.stage.patn.prodv.net
TCRM := bus.stage.tcrm.prodv.net

HOST=none
USER=none
HOME_DIR=$(shell ssh $(USER)@$(HOST) pwd)
PATH_EXCEL_TEST=tmp/excelTest

ifneq ($(BUILD_ENV), prod)
TAG_BUILD_PREF:=$(BUILD_ENV)
endif

info: ## Info about
	@echo TAG_FROM=$(TAG_FROM)
	@echo HOST=$(HOST)
	@echo USER=$(USER)
	@echo HOME_DIR=$(HOME_DIR)

upload-app:
	ssh $(USER)@$(HOST) mkdir -p /home/$(USER)/$(PATH_EXCEL_TEST)
	scp target/ParseXls.jar $(USER)@$(HOST):$(PATH_EXCEL_TEST)

upload-data: upload-app
	scp small.xls $(USER)@$(HOST):$(PATH_EXCEL_TEST)
	scp big.xls $(USER)@$(HOST):$(PATH_EXCEL_TEST)
	scp big.xlsx $(USER)@$(HOST):$(PATH_EXCEL_TEST)

ch_var_231:
	$(eval HOST=$(B231))
	$(eval USER=auto3n)

ch_var_prod:
	$(eval HOST=$(S4))
	$(eval USER=auto3n)

ch_var_loyal:
	$(eval HOST=$(LOYAL))
	$(eval USER=pg-bus)

deploy-231: ch_var_231 upload-data ## deploy 231
deploy-loyal: ch_var_loyal upload-data ## deploy loyal
deploy-prod: ch_var_prod upload-data ## deploy prod

deploy-231-fast: ch_var_231 upload-app ## deploy 231 fast
deploy-loyal-fast: ch_var_loyal upload-app ## deploy loyal fast
deploy-prod-fast: ch_var_prod upload-app ## deploy prod fast
