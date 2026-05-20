import { Button } from '@base-ui/react/button'
import { ScrollArea } from '@base-ui/react/scroll-area'
import { ArrowUp, Square } from 'lucide-react'
import {
  AGENT_COMPOSER_MAX_HEIGHT,
  AGENT_COMPOSER_MIN_HEIGHT,
  agentPanelComposerFrameClass,
  agentPanelComposerScrollContentClass,
  agentPanelComposerScrollRootClass,
  agentPanelComposerScrollViewportClass,
  agentPanelComposerSendButtonClass,
  agentPanelComposerTextareaClass,
  agentPanelScrollAreaScrollbarClass,
  agentPanelScrollAreaThumbClass,
} from './workbook-agent-panel-primitives.js'
import { useAutoSizingTextarea } from './use-autosizing-textarea.js'

export interface WorkbookAgentComposerState {
  readonly sendAriaLabel: 'Send message' | 'Stop'
  readonly isSendDisabled: boolean
}

export function getWorkbookAgentComposerState(input: {
  readonly canInterruptTurn: boolean
  readonly draft: string
  readonly isLoading: boolean
  readonly isRunning: boolean
}): WorkbookAgentComposerState {
  return {
    sendAriaLabel: input.isRunning ? 'Stop' : 'Send message',
    isSendDisabled: input.isRunning ? !input.canInterruptTurn : input.draft.trim().length === 0 || input.isLoading,
  }
}

export function shouldSubmitWorkbookAgentComposerKey(input: {
  readonly isRunning: boolean
  readonly key: string
  readonly shiftKey: boolean
  readonly isComposing: boolean
}): boolean {
  return !input.isRunning && input.key === 'Enter' && !input.shiftKey && !input.isComposing
}

export function WorkbookAgentComposer(props: {
  readonly canInterruptTurn: boolean
  readonly draft: string
  readonly isLoading: boolean
  readonly isRunning: boolean
  readonly onDraftChange: (value: string) => void
  readonly onInterrupt: () => void
  readonly onSubmit: () => void
}) {
  const { textareaRef, viewportRef } = useAutoSizingTextarea({
    value: props.draft,
    minHeight: AGENT_COMPOSER_MIN_HEIGHT,
    maxHeight: AGENT_COMPOSER_MAX_HEIGHT,
  })
  const composerState = getWorkbookAgentComposerState({
    canInterruptTurn: props.canInterruptTurn,
    draft: props.draft,
    isLoading: props.isLoading,
    isRunning: props.isRunning,
  })

  return (
    <form
      onSubmit={(event) => {
        event.preventDefault()
        props.onSubmit()
      }}
    >
      <label className="sr-only" htmlFor="workbook-agent-input">
        Ask the workbook assistant
      </label>
      <div className={agentPanelComposerFrameClass()}>
        <ScrollArea.Root className={agentPanelComposerScrollRootClass()}>
          <ScrollArea.Viewport
            ref={viewportRef}
            className={agentPanelComposerScrollViewportClass()}
            data-testid="workbook-agent-input-viewport"
          >
            <ScrollArea.Content className={agentPanelComposerScrollContentClass()}>
              <textarea
                ref={textareaRef}
                id="workbook-agent-input"
                className={agentPanelComposerTextareaClass()}
                data-testid="workbook-agent-input"
                placeholder="Ask the workbook assistant"
                value={props.draft}
                onChange={(event) => {
                  props.onDraftChange(event.target.value)
                }}
                onKeyDown={(event) => {
                  if (
                    !shouldSubmitWorkbookAgentComposerKey({
                      isRunning: props.isRunning,
                      key: event.key,
                      shiftKey: event.shiftKey,
                      isComposing: event.nativeEvent.isComposing,
                    })
                  ) {
                    return
                  }
                  event.preventDefault()
                  props.onSubmit()
                }}
              />
            </ScrollArea.Content>
          </ScrollArea.Viewport>
          <ScrollArea.Scrollbar className={agentPanelScrollAreaScrollbarClass()} orientation="vertical">
            <ScrollArea.Thumb className={agentPanelScrollAreaThumbClass()} />
          </ScrollArea.Scrollbar>
        </ScrollArea.Root>
        <Button
          aria-label={composerState.sendAriaLabel}
          className={agentPanelComposerSendButtonClass()}
          data-testid="workbook-agent-send"
          disabled={composerState.isSendDisabled}
          type="button"
          onClick={() => {
            if (props.isRunning) {
              if (!props.canInterruptTurn) {
                return
              }
              props.onInterrupt()
              return
            }
            props.onSubmit()
          }}
        >
          {props.isRunning ? <StopIcon /> : <SendArrowIcon />}
        </Button>
      </div>
    </form>
  )
}

function SendArrowIcon() {
  return <ArrowUp aria-hidden="true" className="size-5" strokeWidth={1.9} />
}

function StopIcon() {
  return <Square aria-hidden="true" className="size-4 fill-current" strokeWidth={0} />
}
